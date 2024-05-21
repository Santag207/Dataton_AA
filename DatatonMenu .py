import sys

class DatatonMenu:
    def __init__(self):
        self.options = {
            "1": self.run_etapa1,
            "2": self.run_etapa2,
            "3": self.run_completo,
            "4": self.exit
        }

    def display_menu(self):
        print("""
        Datatón Bancolombia 2023
        Menú de Opciones:
        1. Ejecutar Etapa 1
        2. Ejecutar Etapa 2
        3. Ejecutar Código Completo
        4. Salir
        """)

    def run(self):
        while True:
            self.display_menu()
            choice = input("Selecciona una opción: ")
            action = self.options.get(choice)
            if action:
                action()
            else:
                print(f"{choice} no es una opción válida")

    def run_etapa1(self):
        print("Ejecutando Etapa 1...")
        from etapa1 import run_etapa1
        run_etapa1()

    def run_etapa2(self):
        print("Ejecutando Etapa 2...")
        from etapa2 import run_etapa2
        run_etapa2()

    def run_completo(self):
        print("Ejecutando Código Completo...")
        from completo import run_completo
        run_completo()

    def exit(self):
        print("Saliendo...")
        sys.exit(0)

if __name__ == "__main__":
    menu = DatatonMenu()
    menu.run()
