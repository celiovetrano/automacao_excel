from models.cadastro_funcionario import CadastroFuncionario
from models.salvar_planilhas_funcionario import SalvarPlanilhasFuncionarios
from models.inicializacao import Inicializacao

class MenuPrincipal:

    @staticmethod
    def opcao_selecionada():
        funcionarios = CadastroFuncionario.carregar_funcionarios()

        while True: 
            opcao = input("Escolha uma opção: ").strip()

            if opcao == '1':
                CadastroFuncionario.adicionar_funcionario(funcionarios)
            elif opcao == '2':
                CadastroFuncionario.consultar_funcionarios(funcionarios)
            elif opcao == '3':
                CadastroFuncionario.alterar_funcionario(funcionarios)
            elif opcao == '4':
                CadastroFuncionario.excluir_funcionario(funcionarios)
            elif opcao == '5':
                SalvarPlanilhasFuncionarios.salvar_planilhas()
            elif opcao == '6':
                print("Saindo do programa.")
                break
            else:
                print("Opção inválida. Por favor, escolha uma opção válida.")
                Inicializacao.limpar_tela()
                Inicializacao.exibir_nome_aplicaçao()
                Inicializacao.menu()




