from scheduler_class import Scheduler
import sys

if __name__ == '__main__':
    group = None
    if '--group' in sys.argv:
        try:
            group_index = sys.argv.index('--group') + 1
            group = sys.argv[group_index]
        except (ValueError, IndexError):
            print("Error: --group flag must be followed by a group name.")
            sys.exit(1)

    is_dry_run = '--dry-run' in sys.argv

    scheduler = Scheduler(group=group, dry_run=is_dry_run)
    if scheduler.sheet:
        solution = scheduler.generate_schedule()
        if solution:
            scheduler.create_and_send_offers(solution)
