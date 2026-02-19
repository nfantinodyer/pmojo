"""Shared threading events for pause/stop control across modules."""
import threading

STOP_EVENT = threading.Event()
PAUSE_EVENT = threading.Event()
PAUSE_EVENT.set()  # Start in "running" (not paused) state
