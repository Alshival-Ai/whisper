use anyhow::{Context, Result};
use clap::Parser;
use serde_json::json;
use std::collections::VecDeque;
use std::fs::{File, OpenOptions};
use std::io::Write;
use std::path::Path;
use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::mpsc::{self, Receiver, SyncSender};
use std::sync::{Arc, Mutex};
use std::thread;
use std::time::{Duration, SystemTime, UNIX_EPOCH};
use wasapi::{
    initialize_mta, DeviceEnumerator, Direction, SampleType, StreamMode, WaveFormat,
};
use whisper_rs::{FullParams, SamplingStrategy, WhisperContext, WhisperContextParameters};

#[derive(Parser, Debug)]
#[command(name = "live_meeting_transcriber_rust")]
struct Args {
    #[arg(long, env = "WHISPER_GGML_MODEL_PATH", default_value = "models/ggml-medium.bin")]
    model_path: String,

    #[arg(long, default_value = "cuda", value_parser = ["cuda", "cpu"])]
    device: String,

    #[arg(long, default_value_t = 16000)]
    sample_rate: u32,

    #[arg(long, default_value_t = 6.0)]
    chunk_seconds: f32,

    #[arg(long, default_value_t = 120)]
    block_milliseconds: u32,

    #[arg(long, default_value = "en")]
    language: String,

    #[arg(long, default_value = "")]
    log_file: String,

    #[arg(long)]
    verbose_chunk_log: bool,

    #[arg(long)]
    disable_mic: bool,

    #[arg(long)]
    disable_loopback: bool,

    #[arg(long, default_value_t = 4)]
    threads: i32,

    #[arg(long, default_value_t = 0.003)]
    silence_threshold: f32,
}

#[derive(Clone)]
struct Logger {
    file: Option<Arc<Mutex<File>>>,
}

impl Logger {
    fn new(path: &str) -> Result<Self> {
        if path.trim().is_empty() {
            return Ok(Self { file: None });
        }

        let path_obj = Path::new(path);
        if let Some(parent) = path_obj.parent() {
            if !parent.as_os_str().is_empty() {
                std::fs::create_dir_all(parent)
                    .with_context(|| format!("Failed to create log directory {:?}", parent))?;
            }
        }

        let file = OpenOptions::new()
            .create(true)
            .append(true)
            .open(path_obj)
            .with_context(|| format!("Failed to open log file {}", path))?;

        Ok(Self {
            file: Some(Arc::new(Mutex::new(file))),
        })
    }

    fn log(&self, message: &str) {
        if let Some(file) = &self.file {
            let ts = now_stamp();
            let line = format!("{} [rust-worker] [{}] {}\n", ts, thread_name(), message);
            if let Ok(mut guard) = file.lock() {
                let _ = guard.write_all(line.as_bytes());
            }
        }
    }
}

#[derive(Clone, Copy, Debug)]
enum SourceKind {
    Microphone,
    Loopback,
}

impl SourceKind {
    fn as_str(self) -> &'static str {
        match self {
            SourceKind::Microphone => "microphone",
            SourceKind::Loopback => "loopback",
        }
    }
}

enum CaptureMessage {
    Samples(SourceKind, Vec<f32>),
    Error(SourceKind, String),
    End(SourceKind),
}

fn now_stamp() -> String {
    let now = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .unwrap_or_else(|_| Duration::from_secs(0));
    format!("{}.{:03}", now.as_secs(), now.subsec_millis())
}

fn thread_name() -> String {
    thread::current().name().unwrap_or("main").to_string()
}

fn emit(value: serde_json::Value, logger: &Logger) {
    let line = value.to_string();
    println!("{}", line);
    logger.log(&format!("emit: {}", line));
}

fn emit_status(message: &str, logger: &Logger) {
    emit(json!({ "event": "status", "message": message }), logger);
}

fn emit_error(message: &str, logger: &Logger) {
    emit(json!({ "event": "error", "message": message }), logger);
}

fn emit_transcript(text: &str, logger: &Logger) {
    emit(json!({ "event": "transcript", "text": text }), logger);
}

fn capture_loop(
    source: SourceKind,
    device_direction: Direction,
    sample_rate: u32,
    block_milliseconds: u32,
    tx: SyncSender<CaptureMessage>,
    running: Arc<AtomicBool>,
    logger: Logger,
) -> Result<()> {
    initialize_mta()
        .ok()
        .context("initialize_mta failed in capture thread")?;

    let enumerator = DeviceEnumerator::new().context("DeviceEnumerator::new failed")?;
    let device = enumerator
        .get_default_device(&device_direction)
        .with_context(|| format!("No default {:?} device", device_direction))?;
    let device_name = device
        .get_friendlyname()
        .unwrap_or_else(|_| "Unknown device".to_string());

    let mut audio_client = device
        .get_iaudioclient()
        .context("get_iaudioclient failed")?;

    let desired_format = WaveFormat::new(
        32,
        32,
        &SampleType::Float,
        sample_rate as usize,
        1,
        None,
    );
    let blockalign = desired_format.get_blockalign() as usize;

    let (_, min_period) = audio_client
        .get_device_period()
        .context("get_device_period failed")?;

    let requested_hns = (block_milliseconds as i64 * 10_000).max(min_period);
    let mode = StreamMode::EventsShared {
        autoconvert: true,
        buffer_duration_hns: requested_hns,
    };

    audio_client
        .initialize_client(&desired_format, &Direction::Capture, &mode)
        .with_context(|| {
            format!(
                "initialize_client failed for source={} on device {}",
                source.as_str(),
                device_name
            )
        })?;

    let h_event = audio_client
        .set_get_eventhandle()
        .context("set_get_eventhandle failed")?;

    let capture_client = audio_client
        .get_audiocaptureclient()
        .context("get_audiocaptureclient failed")?;

    let mut sample_queue: VecDeque<u8> = VecDeque::with_capacity(blockalign * 8192);

    audio_client.start_stream().context("start_stream failed")?;
    logger.log(&format!(
        "capture started source={} device_direction={:?} device='{}' sample_rate={} blockalign={}",
        source.as_str(),
        device_direction,
        device_name,
        sample_rate,
        blockalign
    ));

    while running.load(Ordering::Relaxed) {
        capture_client
            .read_from_device_to_deque(&mut sample_queue)
            .context("read_from_device_to_deque failed")?;

        if blockalign >= 4 && sample_queue.len() >= blockalign {
            let frames = sample_queue.len() / blockalign;
            let mut samples = Vec::with_capacity(frames);

            for _ in 0..frames {
                let b0 = sample_queue.pop_front().unwrap_or_default();
                let b1 = sample_queue.pop_front().unwrap_or_default();
                let b2 = sample_queue.pop_front().unwrap_or_default();
                let b3 = sample_queue.pop_front().unwrap_or_default();
                for _ in 4..blockalign {
                    let _ = sample_queue.pop_front();
                }
                samples.push(f32::from_le_bytes([b0, b1, b2, b3]));
            }

            if !samples.is_empty() {
                tx.send(CaptureMessage::Samples(source, samples))
                    .context("failed to send capture samples")?;
            }
        }

        let _ = h_event.wait_for_event(200);
    }

    audio_client.stop_stream().context("stop_stream failed")?;
    logger.log(&format!("capture stopped source={}", source.as_str()));
    Ok(())
}

fn spawn_capture_thread(
    source: SourceKind,
    device_direction: Direction,
    sample_rate: u32,
    block_milliseconds: u32,
    tx: SyncSender<CaptureMessage>,
    running: Arc<AtomicBool>,
    logger: Logger,
) {
    let thread_name = format!("capture-{}", source.as_str());
    let _ = thread::Builder::new().name(thread_name).spawn(move || {
        if let Err(err) = capture_loop(
            source,
            device_direction,
            sample_rate,
            block_milliseconds,
            tx.clone(),
            running,
            logger.clone(),
        ) {
            let message = format!("{} capture failed: {:#}", source.as_str(), err);
            logger.log(&message);
            let _ = tx.send(CaptureMessage::Error(source, message));
        }

        let _ = tx.send(CaptureMessage::End(source));
    });
}

fn drain_chunk(buffer: &mut Vec<f32>, frames: usize) -> Option<Vec<f32>> {
    if buffer.len() < frames {
        return None;
    }

    Some(buffer.drain(..frames).collect())
}

fn transcribe_chunk(
    state: &mut whisper_rs::WhisperState,
    audio: &[f32],
    language: &str,
    threads: i32,
) -> Result<String> {
    let mut params = FullParams::new(SamplingStrategy::Greedy { best_of: 1 });
    params.set_n_threads(threads);
    params.set_translate(false);
    params.set_no_context(true);
    params.set_language(Some(language));
    params.set_print_special(false);
    params.set_print_progress(false);
    params.set_print_realtime(false);
    params.set_print_timestamps(false);

    state
        .full(params, audio)
        .context("whisper full() inference failed")?;

    let mut out = String::new();
    for seg in state.as_iter() {
        let text = seg
            .to_str_lossy()
            .context("failed to extract segment text")?
            .trim()
            .to_string();
        if !text.is_empty() {
            if !out.is_empty() {
                out.push(' ');
            }
            out.push_str(&text);
        }
    }

    Ok(out)
}

fn resolve_device(args: &Args, logger: &Logger) -> (&'static str, bool) {
    if args.device == "cuda" {
        #[cfg(feature = "cuda")]
        {
            return ("cuda", true);
        }

        #[cfg(not(feature = "cuda"))]
        {
            emit_status(
                "Rust worker built without CUDA support. Falling back to CPU.",
                logger,
            );
            return ("cpu", false);
        }
    }

    ("cpu", false)
}

fn run(args: Args) -> Result<i32> {
    let logger = Logger::new(&args.log_file)?;
    logger.log(&format!("args: {:?}", args));

    emit_status("Rust transcriber starting.", &logger);

    let running = Arc::new(AtomicBool::new(true));
    {
        let running_flag = running.clone();
        let logger_clone = logger.clone();
        ctrlc::set_handler(move || {
            running_flag.store(false, Ordering::Relaxed);
            logger_clone.log("ctrl-c received, stopping.");
        })
        .context("failed to set ctrl-c handler")?;
    }

    let (device_label, use_gpu) = resolve_device(&args, &logger);
    emit_status(
        &format!(
            "Loading GGML model {} on {}...",
            args.model_path, device_label
        ),
        &logger,
    );

    let mut ctx_params = WhisperContextParameters::new();
    ctx_params.use_gpu = use_gpu;

    let model_ctx = WhisperContext::new_with_params(&args.model_path, ctx_params)
        .with_context(|| format!("Failed to load model file {}", args.model_path))?;
    let mut state = model_ctx
        .create_state()
        .context("Failed to create whisper state")?;

    emit_status("Model loaded. Live transcription started.", &logger);

    let (tx, rx): (SyncSender<CaptureMessage>, Receiver<CaptureMessage>) = mpsc::sync_channel(64);

    let mut mic_active = false;
    let mut loop_active = false;

    if !args.disable_mic {
        spawn_capture_thread(
            SourceKind::Microphone,
            Direction::Capture,
            args.sample_rate,
            args.block_milliseconds,
            tx.clone(),
            running.clone(),
            logger.clone(),
        );
        mic_active = true;
        emit_status("Microphone source: initializing.", &logger);
    }

    if !args.disable_loopback {
        spawn_capture_thread(
            SourceKind::Loopback,
            Direction::Render,
            args.sample_rate,
            args.block_milliseconds,
            tx.clone(),
            running.clone(),
            logger.clone(),
        );
        loop_active = true;
        emit_status("System audio source: initializing loopback.", &logger);
    }

    if !mic_active && !loop_active {
        emit_error("No sources enabled. Enable microphone and/or loopback.", &logger);
        return Ok(2);
    }

    let chunk_frames = ((args.sample_rate as f32) * args.chunk_seconds).max(1024.0) as usize;

    let mut mic_buffer: Vec<f32> = Vec::with_capacity(chunk_frames * 2);
    let mut loop_buffer: Vec<f32> = Vec::with_capacity(chunk_frames * 2);
    let mut chunk_index: u64 = 0;

    while running.load(Ordering::Relaxed) && (mic_active || loop_active) {
        match rx.recv_timeout(Duration::from_millis(250)) {
            Ok(CaptureMessage::Samples(source, samples)) => match source {
                SourceKind::Microphone => mic_buffer.extend(samples),
                SourceKind::Loopback => loop_buffer.extend(samples),
            },
            Ok(CaptureMessage::Error(source, message)) => {
                emit_error(&message, &logger);
                match source {
                    SourceKind::Microphone => mic_active = false,
                    SourceKind::Loopback => loop_active = false,
                }
            }
            Ok(CaptureMessage::End(source)) => {
                logger.log(&format!("source ended: {}", source.as_str()));
                match source {
                    SourceKind::Microphone => mic_active = false,
                    SourceKind::Loopback => loop_active = false,
                }
            }
            Err(mpsc::RecvTimeoutError::Timeout) => {}
            Err(mpsc::RecvTimeoutError::Disconnected) => {
                emit_error("Capture channel disconnected.", &logger);
                break;
            }
        }

        let mut contributors = 0usize;
        let mut mixed = vec![0.0_f32; chunk_frames];

        if let Some(chunk) = drain_chunk(&mut mic_buffer, chunk_frames) {
            for (idx, sample) in chunk.iter().enumerate() {
                mixed[idx] += *sample;
            }
            contributors += 1;
        }

        if let Some(chunk) = drain_chunk(&mut loop_buffer, chunk_frames) {
            for (idx, sample) in chunk.iter().enumerate() {
                mixed[idx] += *sample;
            }
            contributors += 1;
        }

        if contributors == 0 {
            continue;
        }

        if contributors > 1 {
            let scale = 1.0_f32 / contributors as f32;
            for sample in &mut mixed {
                *sample *= scale;
            }
        }

        chunk_index += 1;
        let peak = mixed
            .iter()
            .fold(0.0_f32, |max_v, v| max_v.max(v.abs()));
        if peak < args.silence_threshold {
            if args.verbose_chunk_log {
                logger.log(&format!(
                    "chunk={} skipped due to low peak {:.6}",
                    chunk_index, peak
                ));
            }
            continue;
        }

        let start = std::time::Instant::now();
        match transcribe_chunk(&mut state, &mixed, &args.language, args.threads) {
            Ok(text) => {
                let elapsed_ms = start.elapsed().as_secs_f64() * 1000.0;
                if args.verbose_chunk_log {
                    logger.log(&format!(
                        "chunk={} contributors={} peak={:.6} infer_ms={:.1} text_len={}",
                        chunk_index,
                        contributors,
                        peak,
                        elapsed_ms,
                        text.len()
                    ));
                }

                if !text.is_empty() {
                    emit_transcript(&text, &logger);
                }
            }
            Err(err) => {
                emit_error(&format!("{}", err), &logger);
            }
        }
    }

    running.store(false, Ordering::Relaxed);
    emit_status("Rust transcriber stopped.", &logger);
    Ok(0)
}

fn main() {
    let args = Args::parse();
    let code = match run(args) {
        Ok(code) => code,
        Err(err) => {
            let line = format!("Fatal error: {:#}", err);
            println!("{}", json!({"event":"error","message":line}));
            1
        }
    };

    std::process::exit(code);
}
