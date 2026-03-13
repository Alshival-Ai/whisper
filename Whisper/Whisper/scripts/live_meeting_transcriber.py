import argparse
import json
import queue
import threading
import time
from dataclasses import dataclass

import numpy as np
import soundcard as sc
import torch
from transformers import pipeline


@dataclass
class CaptureResult:
    frames: np.ndarray
    had_data: bool


class AudioCaptureWorker(threading.Thread):
    def __init__(self, source, sample_rate: int, block_frames: int, label: str) -> None:
        super().__init__(daemon=True)
        self._source = source
        self._sample_rate = sample_rate
        self._block_frames = block_frames
        self._label = label
        self._queue: queue.Queue[np.ndarray] = queue.Queue(maxsize=200)
        self._stop_event = threading.Event()
        self.error: Exception | None = None

    @property
    def has_data(self) -> bool:
        return not self._queue.empty()

    def stop(self) -> None:
        self._stop_event.set()

    def run(self) -> None:
        try:
            with self._source.recorder(samplerate=self._sample_rate, channels=1) as recorder:
                while not self._stop_event.is_set():
                    frame = recorder.record(numframes=self._block_frames)
                    if frame.ndim > 1:
                        frame = frame[:, 0]
                    frame = frame.astype(np.float32, copy=False)
                    try:
                        self._queue.put(frame, timeout=0.25)
                    except queue.Full:
                        # Drop oldest chunk if UI/transcriber is lagging.
                        try:
                            _ = self._queue.get_nowait()
                            self._queue.put_nowait(frame)
                        except queue.Empty:
                            pass
                        except queue.Full:
                            pass
        except Exception as ex:  # pragma: no cover - runtime-specific branch
            self.error = ex

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
    print(json.dumps(message, ensure_ascii=True), flush=True)


def resolve_loopback_source() -> object | None:
    speaker = sc.default_speaker()
    if speaker is None:
        return None

    try:
        return sc.get_microphone(speaker.name, include_loopback=True)
    except Exception:
        try:
            return speaker
        except Exception:
            return None


def build_pipeline(model_name: str, requested_device: str):
    if requested_device == "cuda" and not torch.cuda.is_available():
        emit("status", message="CUDA not available, falling back to CPU.")
        requested_device = "cpu"

    device_index = 0 if requested_device == "cuda" else -1
    dtype = torch.float16 if requested_device == "cuda" else torch.float32

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
    parser.add_argument("--model", default="openai/whisper-large-v3-turbo")
    parser.add_argument("--device", default="cuda", choices=["cuda", "cpu"])
    parser.add_argument("--sample-rate", type=int, default=16000)
    parser.add_argument("--chunk-seconds", type=float, default=6.0)
    parser.add_argument("--block-milliseconds", type=int, default=120)
    parser.add_argument("--language", default="en")
    return parser.parse_args()


def main() -> int:
    args = parse_args()

    mic = sc.default_microphone()
    loopback = resolve_loopback_source()

    if mic is None and loopback is None:
        emit("error", message="No microphone or loopback source available.")
        return 2

    if mic is not None:
        emit("status", message=f"Microphone source: {getattr(mic, 'name', 'default')}.")
    else:
        emit("status", message="Microphone source: unavailable.")

    if loopback is not None:
        emit("status", message=f"System audio source: {getattr(loopback, 'name', 'default loopback')}.")
    else:
        emit("status", message="System audio source: unavailable.")

    asr = build_pipeline(args.model, args.device)

    block_frames = max(256, int(args.sample_rate * (args.block_milliseconds / 1000.0)))
    chunk_frames = max(block_frames, int(args.sample_rate * args.chunk_seconds))
    chunk_timeout = max(args.chunk_seconds * 1.7, 3.0)

    workers: list[AudioCaptureWorker] = []
    if mic is not None:
        workers.append(AudioCaptureWorker(mic, args.sample_rate, block_frames, "microphone"))
    if loopback is not None:
        workers.append(AudioCaptureWorker(loopback, args.sample_rate, block_frames, "loopback"))

    for worker in workers:
        worker.start()

    try:
        while True:
            sources: list[np.ndarray] = []
            had_data = False

            for worker in workers:
                result = worker.read_chunk(chunk_frames, timeout_seconds=chunk_timeout)
                if worker.error is not None:
                    raise worker.error
                if result.had_data:
                    had_data = True
                sources.append(result.frames)

            if not had_data:
                continue

            mixed = np.mean(sources, axis=0).astype(np.float32)
            peak = float(np.max(np.abs(mixed)))
            if peak < 0.003:
                continue

            result = asr(
                {"array": mixed, "sampling_rate": args.sample_rate},
                generate_kwargs={"task": "transcribe", "language": args.language},
            )
            text = (result.get("text") if isinstance(result, dict) else "") or ""
            text = text.strip()
            if text:
                emit("transcript", text=text)
    except KeyboardInterrupt:
        emit("status", message="Transcriber interrupted.")
        return 0
    except Exception as ex:
        emit("error", message=f"{type(ex).__name__}: {ex}")
        return 1
    finally:
        for worker in workers:
            worker.stop()
        for worker in workers:
            worker.join(timeout=2.0)


if __name__ == "__main__":
    raise SystemExit(main())
