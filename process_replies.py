from scheduler_class import Scheduler
import sys

import config

if __name__ == '__main__':
    is_dry_run = '--dry-run' in sys.argv
    scheduler = Scheduler(dry_run=is_dry_run)

    if scheduler.sheet:
        accepted, declined = scheduler.process_email_replies()
        final_sandbox_df = scheduler.finalize_schedule(scheduler.sandbox_df)
        scheduler.redistribute_tokens(final_sandbox_df)

        # Read requester names from metadata sheet to add context to summary
        metadata_ws = scheduler.sheet.worksheet(config.METADATA_TAB)
        requester_names = metadata_ws.acell(config.METADATA_CELL_REQUESTERS).value
        scheduler.send_hr_summary(accepted, declined, requester_names)
