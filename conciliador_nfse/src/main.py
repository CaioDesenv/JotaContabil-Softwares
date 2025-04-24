from controller import ReconciliationController
from view import ReconciliationView

def main():
    controller = ReconciliationController()
    view = ReconciliationView(controller)
    view.iniciar()

if __name__ == '__main__':
    main()
