from scheduler_class import Scheduler
import sys

if __name__ == '__main__':
    is_dry_run = '--dry-run' in sys.argv
    scheduler = Scheduler(dry_run=is_dry_run)

    if scheduler.sheet:
        accepted, declined = scheduler.process_email_replies()
        final_sandbox_df = scheduler.finalize_schedule(scheduler.sandbox_df)
        scheduler.redistribute_tokens(final_sandbox_df)
        scheduler.send_hr_summary(accepted, declined)
