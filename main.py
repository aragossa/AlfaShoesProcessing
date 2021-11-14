import datetime
import os
import sys
from Application.Application import ReportManager

def main():
    args = sys.argv
    app = ReportManager(args=args)
    app.run()


if __name__ == "__main__":
    # os.remove('logs/output.log')
    datetime_start = datetime.datetime.now()
    main()
    delta = datetime.datetime.now() - datetime_start
    print(f'Whole time: {delta}')
    print('Задача выполнена можно закрыть окно')

