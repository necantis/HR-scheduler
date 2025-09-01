from scheduler_class import Scheduler
import sys

if __name__ == '__main__':
    group = None
    if len(sys.argv) > 1 and sys.argv[1] == '--group':
        group = sys.argv[2]

    scheduler = Scheduler(group=group)
    if scheduler.sheet:
        solution = scheduler.generate_schedule()
        if solution:
            scheduler.create_and_send_offers(solution)
