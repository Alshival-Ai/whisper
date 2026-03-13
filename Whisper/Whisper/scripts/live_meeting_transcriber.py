import argparse
import json
import os
import platform
import queue
import sys
import threading
import time
import traceback
from dataclasses import dataclass

import numpy as np

try:
    import sounddevice as sd
    import torch
    from transformers import pipeline
except ModuleNotFoundError as import_error:  # pragma: no cover - environment specific
    sd = None
    torch = None
    pipeline = None
    _IMPORT_ERROR = import_error
else:
    _IMPORT_ERROR = None


class DebugFileLogger:
    def __init__(self) -> None:
        self._path: str | None = None
        self._lock = threading.Lock()

    def configure(self, log_path: str) -> None:
        resolved = os.path.abspath(log_path)
        parent = os.path.dirname(resolved)
        if parent:
            os.makedirs(parent, exist_ok=True)
        self._path = resolved
        self.log(f"Logger configured. path={resolved}")

    @property
    def path(self) -> str | None:
        return self._path

    def log(self, message: str) -> None:
        if self._path is None:
            return

        timestamp = time.strftime("%Y-%m-%d %H:%M:%S", time.localtime())
        milliseconds = int((time.time() % 1) * 1000)
        line = f"{timestamp}.{milliseconds:03d} [worker] [{threading.current_thread().name}] {message}\n"
        with self._lock:
            with open(self._path, "a", encoding="utf-8") as file_handle:
                file_handle.write(line)


LOGGER = DebugFileLogger()


def log_debug(message: str) -> None:
    try:
        LOGGER.log(message)
    except Exception:
        # Debug logging should never break the transcriber process.
        pass


@dataclass
class CaptureResult:
    frames: np.ndarray
    had_data: bool


class AudioCaptureWorker(threading.Thread):
    def __init__(
        self,
        device_index: int,
        target_sample_rate: int,
        block_milliseconds: int,
        label: str,
        *,
        loopback: bool,
    ) -> None:
        super().__init__(daemon=True)
        self._device_index = device_index
        self._target_sample_rate = target_sample_rate
        self._block_milliseconds = block_milliseconds
        self._label = label
        self._loopback = loopback
        self._queue: queue.Queue[np.ndarray] = queue.Queue(maxsize=200)
        self._stop_event = threading.Event()
        self._started_at = time.monotonic()
        self._stream_opened_event = threading.Event()
        self._first_frame_event = threading.Event()
        self.error: Exception | None = None

    def stop(self) -> None:
        self._stop_event.set()

    @property
    def label(self) -> str:
        return self._label

    @property
    def stream_opened(self) -> bool:
        return self._stream_opened_event.is_set()

    @property
    def first_frame_received(self) -> bool:
        return self._first_frame_event.is_set()

    @property
    def seconds_since_start(self) -> float:
        return time.monotonic() - self._started_at

    def run(self) -> None:
        log_debug(
            f"AudioCaptureWorker starting. label={self._label}, device_index={self._device_index}, "
            f"target_sample_rate={self._target_sample_rate}, block_milliseconds={self._block_milliseconds}, loopback={self._loopback}"
        )
        try:
            device_info = sd.query_devices(self._device_index)
            source_sample_rate = int(round(float(device_info.get("default_samplerate") or self._target_sample_rate)))
            if source_sample_rate <= 0:
                source_sample_rate = self._target_sample_rate

            max_input_channels = int(device_info.get("max_input_channels", 0) or 0)
            channels = max(1, min(2, max_input_channels if max_input_channels > 0 else 1))
            extra_settings = None

            stream_blocksize = max(256, int(source_sample_rate * (self._block_milliseconds / 1000.0)))
            log_debug(
                f"AudioCaptureWorker stream config. label={self._label}, device='{device_info.get('name', self._device_index)}', "
                f"source_sample_rate={source_sample_rate}, stream_blocksize={stream_blocksize}, channels={channels}, "
                f"loopback={self._loopback}"
            )
            first_callback_logged = False

            def callback(indata, frames, _time_info, status) -> None:
                nonlocal first_callback_logged
                if status:
                    log_debug(f"sounddevice status for {self._label}: {status}")

                if not first_callback_logged:
                    first_callback_logged = True
                    self._first_frame_event.set()
                    log_debug(f"AudioCaptureWorker first callback. label={self._label}, frames={frames}")

                frame = np.asarray(indata, dtype=np.float32)
                if frame.ndim == 2:
                    mono = np.mean(frame, axis=1, dtype=np.float32)
                else:
                    mono = frame.astype(np.float32, copy=False)

                if source_sample_rate != self._target_sample_rate:
                    mono = resample_linear(mono, source_sample_rate, self._target_sample_rate)

                if mono.size == 0:
                    return

                try:
                    self._queue.put_nowait(mono)
                except queue.Full:
                    log_debug(f"Queue full for source={self._label}. Dropping oldest frame block.")
                    try:
                        _ = self._queue.get_nowait()
                        self._queue.put_nowait(mono)
                    except queue.Empty:
                        pass
                    except queue.Full:
                        pass

            with sd.InputStream(
                samplerate=source_sample_rate,
                blocksize=stream_blocksize,
                device=self._device_index,
                channels=channels,
                dtype=np.float32,
                callback=callback,
                extra_settings=extra_settings,
            ):
                self._stream_opened_event.set()
                log_debug(f"AudioCaptureWorker stream opened. label={self._label}")
                while not self._stop_event.is_set():
                    time.sleep(0.05)
        except Exception as ex:  # pragma: no cover - runtime-specific branch
            self.error = ex
            log_debug(f"AudioCaptureWorker exception. label={self._label}, error={type(ex).__name__}: {ex}")
            log_debug(traceback.format_exc())
        finally:
            log_debug(f"AudioCaptureWorker exiting. label={self._label}")

    def read_chunk(self, target_frames: int, timeout_seconds: float) -> CaptureResult:
        deadline = time.time() + timeout_seconds
        parts: list[np.ndarray] = []
        collected = 0

        while collected < target_frames and time.time() < deadline and not self._stop_event.is_set():
            remaining = deadline - time.time()
            wait_seconds = min(max(remaining, 0.0), 0.25)
            try:
                frame = self._queue.get(timeout=wait_seconds)
            except queue.Empty:
                continue

            parts.append(frame)
            collected += frame.shape[0]

        if not parts:
            return CaptureResult(frames=np.zeros(target_frames, dtype=np.float32), had_data=False)

        concatenated = np.concatenate(parts)
        if concatenated.shape[0] < target_frames:
            concatenated = np.pad(concatenated, (0, target_frames - concatenated.shape[0]))
        elif concatenated.shape[0] > target_frames:
            concatenated = concatenated[:target_frames]

        return CaptureResult(frames=concatenated, had_data=True)


def emit(event_name: str, **payload) -> None:
    message = {"event": event_name}
    message.update(payload)
    serialized = json.dumps(message, ensure_ascii=True)
    print(serialized, flush=True)
    log_debug(f"emit: {serialized}")


def get_default_device_index(is_input: bool) -> int | None:
    if sd is None:
        return None

    key = "default_input_device" if is_input else "default_output_device"
    try:
        for hostapi in sd.query_hostapis():
            name = str(hostapi.get("name", ""))
            if "WASAPI" in name.upper():
                candidate = int(hostapi.get(key, -1))
                if candidate >= 0:
                    return candidate
    except Exception as ex:
        log_debug(f"WASAPI hostapi lookup failed: {type(ex).__name__}: {ex}")

    try:
        default_pair = sd.default.device
        if isinstance(default_pair, (tuple, list)) and len(default_pair) >= 2:
            candidate = int(default_pair[0] if is_input else default_pair[1])
            if candidate >= 0:
                return candidate
    except Exception as ex:
        log_debug(f"Fallback default device lookup failed: {type(ex).__name__}: {ex}")

    return None


def get_loopback_device_index() -> int | None:
    if sd is None:
        return None

    preferred_terms = ("stereo mix", "loopback", "what u hear", "wave out mix")
    candidates: list[tuple[int, int]] = []

    try:
        for entry in sd.query_devices():
            max_input_channels = int(entry.get("max_input_channels", 0) or 0)
            if max_input_channels <= 0:
                continue

            name = str(entry.get("name", ""))
            lowered = name.lower()
            if not any(term in lowered for term in preferred_terms):
                continue

            hostapi_index = int(entry.get("hostapi", -1) or -1)
            hostapi_name = ""
            if hostapi_index >= 0:
                try:
                    hostapi_name = str(sd.query_hostapis(hostapi_index).get("name", "")).lower()
                except Exception:
                    hostapi_name = ""

            score = 100
            if "wasapi" in hostapi_name:
                score += 10
            elif "wdm-ks" in hostapi_name:
                score += 8

            score += min(max_input_channels, 2)
            candidates.append((score, int(entry["index"])))
    except Exception as ex:
        log_debug(f"Loopback device lookup failed: {type(ex).__name__}: {ex}")
        return None

    if not candidates:
        return None

    candidates.sort(reverse=True)
    return candidates[0][1]


def get_device_name(device_index: int | None) -> str:
    if sd is None or device_index is None:
        return "unavailable"

    try:
        info = sd.query_devices(device_index)
        return str(info.get("name", f"device {device_index}"))
    except Exception as ex:
        log_debug(f"Failed to query device name for index={device_index}: {type(ex).__name__}: {ex}")
        return f"device {device_index}"


def resample_linear(samples: np.ndarray, source_rate: int, target_rate: int) -> np.ndarray:
    if source_rate <= 0 or target_rate <= 0 or source_rate == target_rate or samples.size == 0:
        return samples.astype(np.float32, copy=False)

    source_len = samples.shape[0]
    target_len = max(1, int(round(source_len * target_rate / source_rate)))
    if source_len == target_len:
        return samples.astype(np.float32, copy=False)

    source_positions = np.linspace(0.0, 1.0, num=source_len, endpoint=False, dtype=np.float64)
    target_positions = np.linspace(0.0, 1.0, num=target_len, endpoint=False, dtype=np.float64)
    return np.interp(target_positions, source_positions, samples).astype(np.float32, copy=False)


def build_pipeline(model_name: str, requested_device: str):
    if requested_device == "cuda" and not torch.cuda.is_available():
        emit("status", message="CUDA not available, falling back to CPU.")
        requested_device = "cpu"

    device_index = 0 if requested_device == "cuda" else -1
    dtype = torch.float16 if requested_device == "cuda" else torch.float32
    log_debug(
        f"Pipeline build requested. model={model_name}, device={requested_device}, torch_dtype={dtype}, device_index={device_index}"
    )

    if requested_device == "cuda":
        try:
            device_name = torch.cuda.get_device_name(0)
            log_debug(f"CUDA device: {device_name}")
        except Exception as ex:
            log_debug(f"Unable to query CUDA device name: {type(ex).__name__}: {ex}")

    emit("status", message=f"Loading model {model_name} on {requested_device}...")
    asr = pipeline(
        task="automatic-speech-recognition",
        model=model_name,
        torch_dtype=dtype,
        device=device_index,
    )
    emit("status", message="Model loaded. Live transcription started.")
    return asr


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Capture mic + loopback audio and stream Whisper transcripts.")
    parser.add_argument("--model", default="openai/whisper-medium")
    parser.add_argument("--device", default="cuda", choices=["cuda", "cpu"])
    parser.add_argument("--sample-rate", type=int, default=16000)
    parser.add_argument("--chunk-seconds", type=float, default=6.0)
    parser.add_argument("--block-milliseconds", type=int, default=120)
    parser.add_argument("--language", default="en")
    parser.add_argument("--log-file", default="")
    parser.add_argument("--verbose-chunk-log", action="store_true")
    parser.add_argument("--disable-mic", action="store_true")
    parser.add_argument("--disable-loopback", action="store_true")
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    if args.log_file:
        LOGGER.configure(args.log_file)

    log_debug("Transcriber process starting.")
    log_debug(f"Python version: {sys.version}")
    log_debug(f"Platform: {platform.platform()}")
    log_debug(f"Python executable: {sys.executable}")
    log_debug(f"Args: {args}")

    if "WindowsApps" in sys.executable and "PythonSoftwareFoundation" in sys.executable:
        emit(
            "status",
            message=(
                "Windows Store Python detected. If device capture fails, set WHISPER_PYTHON to a standard python.org install."
            ),
        )

    if _IMPORT_ERROR is not None:
        missing_package = getattr(_IMPORT_ERROR, "name", "unknown")
        log_debug(f"Import error encountered: {type(_IMPORT_ERROR).__name__}: {_IMPORT_ERROR}")
        emit(
            "error",
            message=(
                f"Missing Python package '{missing_package}'. "
                "Run: python -m pip install -r Whisper\\Whisper\\scripts\\requirements-live-transcriber.txt"
            ),
        )
        return 3

    mic_device_index = None if args.disable_mic else get_default_device_index(is_input=True)
    loopback_device_index = None if args.disable_loopback else get_loopback_device_index()

    log_debug(f"Resolved device indexes. mic={mic_device_index}, loopback={loopback_device_index}")

    if mic_device_index is None and loopback_device_index is None:
        log_debug("No microphone or loopback source available.")
        emit("error", message="No microphone or loopback source available.")
        return 2

    if mic_device_index is not None:
        emit("status", message=f"Microphone source: {get_device_name(mic_device_index)}.")
    else:
        emit("status", message="Microphone source: unavailable.")

    if loopback_device_index is not None:
        emit("status", message=f"System audio source: {get_device_name(loopback_device_index)}.")
    else:
        emit("status", message="System audio source: unavailable (no loopback/stereo-mix input found).")

    asr = build_pipeline(args.model, args.device)
    log_debug("Pipeline loaded successfully.")

    block_frames = max(256, int(args.sample_rate * (args.block_milliseconds / 1000.0)))
    chunk_frames = max(block_frames, int(args.sample_rate * args.chunk_seconds))
    chunk_timeout = max(args.chunk_seconds * 1.7, 3.0)
    log_debug(
        f"Audio chunk config. block_frames={block_frames}, chunk_frames={chunk_frames}, timeout={chunk_timeout:.2f}s"
    )

    workers: list[AudioCaptureWorker] = []
    if mic_device_index is not None:
        workers.append(
            AudioCaptureWorker(
                mic_device_index,
                args.sample_rate,
                args.block_milliseconds,
                "microphone",
                loopback=False,
            )
        )
    if loopback_device_index is not None:
        workers.append(
            AudioCaptureWorker(
                loopback_device_index,
                args.sample_rate,
                args.block_milliseconds,
                "loopback",
                loopback=True,
            )
        )

    for worker in workers:
        worker.start()
    log_debug(f"Workers started. count={len(workers)}")

    try:
        chunk_index = 0
        stream_open_timeout = max(6.0, args.chunk_seconds + 2.0)
        first_frame_timeout = max(10.0, args.chunk_seconds * 2.0)
        while True:
            chunk_index += 1
            sources: list[np.ndarray] = []
            had_data = False

            active_workers: list[AudioCaptureWorker] = []

            for worker in workers:
                if worker.error is None and not worker.stream_opened and worker.seconds_since_start > stream_open_timeout:
                    worker.error = TimeoutError(
                        f"{worker.label} capture stream did not open within {stream_open_timeout:.1f}s."
                    )

                if (
                    worker.error is None
                    and worker.stream_opened
                    and not worker.first_frame_received
                    and worker.seconds_since_start > first_frame_timeout
                ):
                    worker.error = TimeoutError(
                        f"{worker.label} capture stream opened but no audio frames arrived within {first_frame_timeout:.1f}s."
                    )

                if worker.error is not None:
                    emit(
                        "error",
                        message=f"{worker.label} capture failed: {type(worker.error).__name__}: {worker.error}",
                    )
                    log_debug(
                        f"Disabling failed source. label={worker.label}, error={type(worker.error).__name__}: {worker.error}"
                    )
                    continue

                result = worker.read_chunk(chunk_frames, timeout_seconds=chunk_timeout)
                if worker.error is not None:
                    emit(
                        "error",
                        message=f"{worker.label} capture failed: {type(worker.error).__name__}: {worker.error}",
                    )
                    log_debug(
                        f"Disabling failed source after read. label={worker.label}, error={type(worker.error).__name__}: {worker.error}"
                    )
                    continue

                active_workers.append(worker)
                if result.had_data:
                    had_data = True
                sources.append(result.frames)

            workers = active_workers

            if not workers:
                emit("error", message="All audio sources failed.")
                return 1

            if not had_data:
                if args.verbose_chunk_log:
                    log_debug(f"chunk={chunk_index} skipped: no source data.")
                continue

            mixed = np.mean(sources, axis=0).astype(np.float32)
            peak = float(np.max(np.abs(mixed)))
            if peak < 0.003:
                if args.verbose_chunk_log:
                    log_debug(f"chunk={chunk_index} skipped: below peak threshold. peak={peak:.6f}")
                continue

            infer_start = time.perf_counter()
            result = asr(
                {"array": mixed, "sampling_rate": args.sample_rate},
                generate_kwargs={"task": "transcribe", "language": args.language},
            )
            infer_ms = (time.perf_counter() - infer_start) * 1000.0
            text = (result.get("text") if isinstance(result, dict) else "") or ""
            text = text.strip()

            if args.verbose_chunk_log:
                log_debug(f"chunk={chunk_index} peak={peak:.6f} infer_ms={infer_ms:.1f} text_len={len(text)}")

            if text:
                log_debug(f"chunk={chunk_index} transcript={text}")
                emit("transcript", text=text)
    except KeyboardInterrupt:
        log_debug("KeyboardInterrupt received.")
        emit("status", message="Transcriber interrupted.")
        return 0
    except Exception as ex:
        log_debug(f"Fatal transcriber exception: {type(ex).__name__}: {ex}")
        log_debug(traceback.format_exc())
        emit("error", message=f"{type(ex).__name__}: {ex}")
        return 1
    finally:
        log_debug("Stopping workers.")
        for worker in workers:
            worker.stop()
        for worker in workers:
            worker.join(timeout=2.0)
        log_debug("Workers stopped. Transcriber process exiting.")


if __name__ == "__main__":
    raise SystemExit(main())
