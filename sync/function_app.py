import asyncio
import logging
import os

import azure.functions as func

from main import main as run_sync

os.environ.setdefault("TIMER_SCHEDULE", "0 0 2 * * *")


app = func.FunctionApp()


@app.function_name(name="sharepoint_sync_timer")
@app.timer_trigger(
    schedule="%TIMER_SCHEDULE%",
    arg_name="timer",
    run_on_startup=False,
    use_monitor=True,
)
def sharepoint_sync_timer(timer: func.TimerRequest) -> None:
    logging.info("Timer trigger received for SharePoint sync")

    exit_code = asyncio.run(run_sync())
    if exit_code != 0:
        raise RuntimeError(f"SharePoint sync failed with exit code {exit_code}")

    logging.info("SharePoint sync completed successfully")
