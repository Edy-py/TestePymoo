from EnergyPlusKernel import *



def clean_idf(caminho_idf: str) -> None:
    """
    Remove o arquivo IDF especificado.

    :param caminho_idf: Caminho para o arquivo do modelo da edificação (.idf) a ser removido
    """
    os.remove(caminho_idf)


def chamar_ep_processar_idf(entrada_idf, ep_install_path, pasta_simulacao, epw_file, material, vidros):
    """
    Função para chamar o EnergyPlus e processar o arquivo IDF com todas as alterações arquitetônicas e de materiais.

    :param entrada_idf: Caminho para o arquivo do modelo da edificação (.idf)

    :return: Hora total de conforto térmico do ambiente simulado
    """

    # Alteração do arquivo IDF
    idf_v = alterar_versao_energyplus_idf(entrada_idf)
    idf_vs = configurar_simulation_control(idf_v)
    clean_idf(idf_v)
    idf_vsb = configurar_building(idf_vs)
    clean_idf(idf_vs)
    idf_vsbt = configurar_timestep(idf_vsb)
    clean_idf(idf_vsb)
    idf_vsbtr = configurar_run_period(idf_vsbt)
    clean_idf(idf_vsbt)
    idf_vsbtrd = configurar_dias_especiais(idf_vsbtr)
    clean_idf(idf_vsbtr)
    idf_vsbtrdm = adicionar_materiais_do_excel(idf_vsbtrd, material)
    clean_idf(idf_vsbtrd)
    idf_vsbtrdmv = adicionar_vidros_do_excel(idf_vsbtrdm, vidros)
    clean_idf(idf_vsbtrdm)
    idf_final = configurar_localizacao_do_epw(idf_vsbtrdmv, epw_file)
    clean_idf(idf_vsbtrdmv)
    idf = salvar_idf_final(entrada_idf, idf_final)
    clean_idf(idf_final)

    # Processar o arquivo IDF
    processar_idf(pasta_simulacao, "Zona 6", ep_install_path, epw_file, idf)