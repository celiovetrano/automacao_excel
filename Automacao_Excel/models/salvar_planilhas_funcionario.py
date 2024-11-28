import os
import re
import calendar
from models.inicializacao import Inicializacao
from models.gerar_planilha import GerarPlanilha
from models.cadastro_funcionario import CadastroFuncionario

class SalvarPlanilhasFuncionarios:
    # Função para gerar planilhas de horas de cada funcionario da lista
    def salvar_planilhas():
        funcionarios = CadastroFuncionario.carregar_funcionarios()
        if not funcionarios:
            print("Nenhum funcionário cadastrado.")
            return

        # Garante que o diretório 'planilhas' exista
        diretorio = "planilhas"
        if not os.path.exists(diretorio):
            os.makedirs(diretorio)

        mes = GerarPlanilha.obter_inteiro_valido("Por favor digite o Mês atual com 2 digitos (MM): ", 1, 12)
        ano = GerarPlanilha.obter_inteiro_valido("Por favor digite o Ano atual com 4 digitos (AAAA): ", 1900, 2100)
        recesso = GerarPlanilha.obter_resposta_sim_nao("Há recesso neste mês? (s/n): ") == 's'
        nome_nova_aba = f"{GerarPlanilha.obter_nome_mes(mes)}_{ano}"
        
        if recesso:
            recesso_inicio = GerarPlanilha.obter_inteiro_valido("Por favor digite o dia de início do recesso (DD): ", 1, calendar.monthrange(ano, mes)[1])
            recesso_fim = GerarPlanilha.obter_inteiro_valido("Por favor digite o dia de fim do recesso (DD): ", recesso_inicio, calendar.monthrange(ano, mes)[1])
        else:
            recesso_inicio, recesso_fim = None, None
                
        for rf, nome in funcionarios.items():
            primeiro_nome = nome.split()[0]
            rf_numeros = re.sub(r'\D', '', rf)  # Remove pontos e barra do RF, mantendo apenas números

            # Caminho do arquivo de origem
            arquivo_origem = "exemplo.xlsm"

            # Gera a Planilha para cada servidor
            GerarPlanilha.editar_aba(arquivo_origem, nome_nova_aba, mes, ano, primeiro_nome, rf_numeros, diretorio, recesso_inicio, recesso_fim)
         
        if input("Pressione 'Tab' para retornar ao menu: ") == '\t':
            Inicializacao.limpar_tela()
            return