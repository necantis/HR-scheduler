from scheduler_class import Scheduler

if __name__ == '__main__':
    scheduler = Scheduler()
    if scheduler.sheet:
        scheduler.process_email_replies()
        final_sandbox_df = scheduler.finalize_schedule(scheduler.sandbox_df)
        scheduler.redistribute_tokens(final_sandbox_df)
