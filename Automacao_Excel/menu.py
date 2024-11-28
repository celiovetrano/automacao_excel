from models.inicializacao import Inicializacao
from models.menu_principal import MenuPrincipal

def main():
    import os
    os.system('cls')
    Inicializacao.exibir_nome_aplicaçao()
    Inicializacao.menu()
    
    MenuPrincipal.opcao_selecionada()


if __name__ == "__main__":
    main()