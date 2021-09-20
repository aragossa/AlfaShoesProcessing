import sys
from Application.Application import ReportManager

def main():
    args = sys.argv
    app = ReportManager(args=args)
    app.run()


if __name__ == "__main__":
    main()
