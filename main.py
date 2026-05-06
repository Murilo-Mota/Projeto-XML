import pandas as pd
import os
import sys
import xmltodict
import ttkbootstrap as tb
import openpyxl
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs import Messagebox, Querybox
from tkinter import filedialog
from ttkbootstrap.tooltip import ToolTip
from pathlib import Path
import threading
from datetime import datetime
import traceback


def pegar(dados, caminho, padrao=None):
    '''
    Função responsável por pegar as informações do XML garantindo que não quebre caso não haja alguma chave
    '''
    for chave in caminho:
        if isinstance(dados, dict) and chave in dados:
            dados = dados[chave]
        else:
            return padrao
    return dados

def regime(dados, caminho):
    '''
    Função para definir o regime tributário do fornecedor.
    '''
    for chave in caminho:
        if isinstance(dados, dict) and chave in dados:
            dados = dados[chave]
            if dados == '1':
                dados = 'Simples'
            elif dados == '2':
                dados = 'SN, Sublimite RB'
            elif dados == '3':
                dados = 'Normal'
            else:
                dados = 'MEI'
        return dados

pasta_xml = ''
pasta_saida = ''
quantidade = 0



def leitor_xml():
    
    """
    Função para percorrer todos os XMLs em uma pasta para processalos e gerar um DataFra,e

    Parametros:
        pasta_xml (str): Caminho para os arquivos.

    Retorno:
        DF dos clientes, Fornecedores e Produtos contidos nos XMLs.
    """
    global pasta_xml
    global pasta_saida
    caminho = Path(pasta_xml)
    xmls = list(caminho.rglob("*.xml"))
    notas = []
    ctes = []
    erros = []
    nfse = []

    for arquivo in xmls:
        

        try:

            with open(arquivo, "r", encoding="utf-8") as file:
                dados = xmltodict.parse(file.read())
                if 'cteProc' in dados:
                    pass #tratativa de CTEs
                elif 'nfeProc' in dados:
                    '''
                    Tratativa para NFe modelo 55
                    '''
                    infNFe = dados['nfeProc']['NFe']['infNFe']
                    ide = infNFe['ide']
                    emit = infNFe['emit']
                    dest = infNFe['dest']
                    det = pegar(infNFe, ['det'], [])

                    if isinstance(det, dict):
                        det = [det]

                    for item in det: 

                        prod = pegar(item,['prod'])
                        icms = pegar(item, ['imposto','ICMS'])
                        difal = pegar(item,['imposto','ICMSUFDest'])
                        ipi = pegar(item, ['imposto','IPI'])
                        pis = pegar(item,['imposto','PIS'])
                        cofins = pegar(item,['imposto','COFINS'])
                        ibscbs = pegar(item, ['imposto','IBSCBS'])
                        cst_icm = next(iter(icms.keys())) if icms else {}



                        notas.append({
                             #ide:
                            'Modelo': pegar(ide, ['mod']),
                            'Numero NF': pegar(ide, ['nNF']),
                            'Serie NF': pegar(ide, ['serie']),
                            'Emissão': pegar(ide, ['dhEmi']),
                            'Finalidade': pegar(ide, ['finNFe']),
                             # Emit:
                            'CNPJ Emitente': pegar(emit, ['CNPJ']) if 'CNPJ' in emit else pegar(emit, ['CPF']),
                            'Razão Social Emitente': pegar(emit, ['xNome']),
                            'IE Emitente': pegar(emit, ['IE'], padrao="ISENTO"),
                            'UF Emitente': pegar(emit, ['enderEmit', 'UF']),
                            'Mun Emitente': pegar(emit,['enderEmit', 'xMun']),
                            'Regime Trib': regime(emit, ['CRT']),

                             #Dest: 
                            'CNPJ Destinatário': pegar(dest, ['CNPJ']) if 'CNPJ' in dest else pegar(dest, ['CPF']),
                            'Razão Social Destinatário': pegar(dest, ['xNome']),
                            'IE Destinatário': pegar(dest, ['IE'], padrao="ISENTO"),
                            'UF Destinatário': pegar(dest, ['enderDest', 'UF']),
                            'Mun Destinatário': pegar(dest,['enderDest', 'xMun']),

                            #prod:
                            'item XML' : item['@nItem'],
                            'Cod.Produto': pegar(item,['prod','cProd']),
                            'Descrição': pegar(item,['prod','xProd']),
                            'NCM': pegar(item,['prod','NCM']),
                            'UM': pegar(item,['prod','uCom']),
                            'CFOP': pegar(item,['prod','CFOP']),
                            'Quantidade': pegar(item,['prod','qCom'], padrao=0),
                            'Valor unitário': pegar(item,['prod','vUnCom'], padrao=0),
                            'Total': pegar(item,['prod','vProd'],0),
                            'cBenef': pegar(item,['prod','cBenef'],''),

                            #tributos:
                            #ICMS:
                            'Origem' :  pegar(icms, [cst_icm, 'orig'],0),
                            'CST ICMS': pegar(icms, [cst_icm, 'CST']) if 'CST' in icms[cst_icm] else pegar(icms, [cst_icm,'CSOSN']),
                            'BC ICMS': pegar(icms, [cst_icm, 'vBC'],0),
                            '% ICMS' : pegar(icms, [cst_icm, 'pICMS'],0),
                            'Valor ICMS' : pegar(icms, [cst_icm, 'vICMS'],0),

                            #ICMS Especifico:
                            #Diferimento:
                            '% Diferimento ICMS': pegar(icms, [cst_icm, 'pDif'],0),
                            '% Redução BC': pegar(icms, [cst_icm, 'pRedBC'],0),

                            #ICMS ST: 

                            'MVA ICMS ST': pegar(icms, [cst_icm, 'pMVAST'],0), 
                            'Base ICMS ST': pegar(icms, [cst_icm, 'vBCST'],0),
                            '% ICMS ST': pegar(icms, [cst_icm, 'pICMSST'],0) if 'pICMSST' in icms[cst_icm] else pegar(icms, [cst_icm, 'vBCSTRet'],0), 
                            'Valor ICMS ST': pegar(icms, [cst_icm, 'vICMSST'],0), 
                            'Valor ICMS Retido': pegar(icms, [cst_icm,'vICMSSTRet'],0),

                            #ICMS Monofásico:
                            'Base ICMS Monofasico' : pegar(icms, [cst_icm,'qBCMono'],0),
                            'AdRem ICMS Monofasico' : pegar(icms, [cst_icm,'adRemICMS'],0),
                            'Valor ICMS Monofasico' : pegar(icms, [cst_icm,'adRemICMS'],0),

                            #IPI:
                            'Enq.IPI': pegar(ipi, ['cEnq'], padrao='999'),

                            'CST IPI': pegar(ipi, ['IPITrib', 'CST'],0) if ipi and 'IPITrib' in ipi else pegar(ipi, ['IPINT', 'CST'],0),
                            'Base IPI': pegar(ipi, ['IPITrib', 'vBC'],0),
                            '% IPI' : pegar(ipi, ['IPITrib','pIPI'],0),
                            'Valor do IPI' : pegar(ipi, ['IPITrib','vIPI'],0),

                            #PIS

                            'CST PIS': pegar(pis, [next(iter(pis.keys())), 'CST'],0),
                            'Base PIS': pegar(pis, [next(iter(pis.keys())), 'vBC'], 0),
                            '% PIS': pegar(pis, [next(iter(pis.keys())), 'pPIS'], 0),
                            'Valor PIS': pegar(pis, [next(iter(pis.keys())), 'vPIS'], 0),

                            #COFINS
                            'CST COFINS': pegar(cofins, [next(iter(cofins.keys())), 'CST']),
                            'Base COFINS': pegar(cofins, [next(iter(cofins.keys())), 'vBC'], 0),
                            '% COFINS': pegar(cofins, [next(iter(cofins.keys())), 'pCOFINS'], 0),
                            'Valor COFINS': pegar(cofins, [next(iter(cofins.keys())), 'vCOFINS'], 0),

                            #IBS e CBS:

                            'CST IBS e CBS': pegar(ibscbs, ['CST'], 'SEM DESTAQUE'),
                            'CCLASTRIB': pegar(ibscbs, ['cClassTrib'], ""),
                            'Base IBS e CBS': pegar(ibscbs, ['gIBSCBS', 'vBC'],0),
                            'Aliq. IBSUF': pegar(ibscbs, ['gIBSCBS', 'gIBSUF', 'pIBSUF'],0),
                            'valor IBSUF': pegar(ibscbs, ['gIBSCBS', 'gIBSUF', 'vIBSUF'],0),
                            'Aliq. IBSMun': pegar(ibscbs, ['gIBSCBS', 'gIBSMun', 'pIBSMun'],0),
                            'valor IBSMun': pegar(ibscbs, ['gIBSCBS', 'gIBSMun', 'vIBSMun'],0),
                            'Aliq. CBS': pegar(ibscbs, ['gIBSCBS', 'gCBS', 'pCBS'],0),
                            'Valor CBS': pegar(ibscbs, ['gIBSCBS', 'gCBS', 'vCBS'],0),

                        })
                    
                elif 'NFSe' in dados:
                    '''
                    Tratativa para NFSE Nacional
                    '''
                    inf_nfse = pegar(dados, ['NFSe', 'infNFSe'])
                    emissor = pegar(inf_nfse, ['emit'])
                    dps = pegar(inf_nfse, ['DPS', 'infDPS'])
                    prestador = pegar(dps, ['prest'])
                    tomador = pegar(dps, ['toma'])
                    servico = pegar(dps, ['serv'])
                    valores = pegar(inf_nfse, ['valores'])
                    trib_fed = pegar(dps, ['valores','trib', 'tribFed'])
                    ibscbs = pegar(inf_nfse, ['IBSCBS'])

                    nfse.append({
                        # ---------------- IDE ----------------
                        'Numero NF': pegar(inf_nfse, ['nNFSe']),
                        'Série': pegar(dps, ['serie'], 'UNI'),
                        'Numero DFSe': pegar(inf_nfse, ['nDFSe']),
                        'Emissão': pegar(inf_nfse, ['dhProc']),
                        'Competencia': pegar(dps, ['dCompet']),
                        #'Competencia': pegar(inf_nfse, ['Competencia']),
                        #'Codigo Verificação': pegar(inf_nfse, ['CodigoVerificacao']),

                        # ---------------- PRESTADOR ----------------
                        'CNPJ Emitente': pegar(emissor, ['CNPJ']) if 'CNPJ' in emissor else pegar(emissor, ['CPF']),
                        #'Inscrição Municipal Emitente': pegar(prestador, ['InscricaoMunicipal']),
                        'Razão Social Emitente': pegar(emissor, ['xNome']),
                        'CNPJ Prestador': pegar(prestador, ['CNPJ']) if 'CNPJ' in prestador else pegar(prestador, ['CPF']),

                        'Município Prestador': pegar(emissor, ['enderNac', 'cMun']),
                        'UF Prestador': pegar(emissor, ['enderNac', 'UF']),
                        'Optante do simples': pegar(prestador, ['regTrib', 'opSimpNac']),
                        'Regime Especial': pegar(prestador, ['regTrib', 'regEspTrib']),

                        #'Regime Especial Tributação': pegar(inf_nfse, ['RegimeEspecialTributacao']),
                        #'Optante Simples Nacional': pegar(inf_nfse, ['OptanteSimplesNacional']),
                        #'Incentivador Cultural': pegar(inf_nfse, ['IncentivadorCultural']),

                        # ---------------- TOMADOR ----------------
                        'CNPJ Tomador': pegar(tomador, ['CNPJ']) or pegar(tomador, ['Cpf']),
                        #'Inscrição Municipal Tomador': pegar(tomador, ['InscricaoMunicipal']),
                        'Razão Social Tomador': pegar(tomador, ['xNome']),

                        'Município Tomador': pegar(tomador, ['end', 'endNac', 'cMun']),

                        # ---------------- SERVIÇO ----------------
                        'Local Prestação': pegar(servico, ['locPrest', 'cLocPrestacao']),
                        'Cod. Local Incidencia': pegar(inf_nfse, ['cLocIncid']),
                        'Local Incidencia': pegar(inf_nfse, ['xLocIncid']),
                        'Código Serviço': pegar(servico, ['cServ', 'cTribNac']),
                        'Descrição Nacional': pegar(inf_nfse, ['xTribNac']),
                        'Código NBS': pegar(servico, ['cServ','cNBS']),
                        'Descrição NBS': pegar(inf_nfse, ['xNBS']),
                        'Cod.Trib. Municipal': pegar(servico, ['cServ', 'cTribMun']),
                        'Descrição Municipal': pegar(inf_nfse, ['xTribMun']),
                        'Descrição do emissor': pegar(servico, ['cServ', 'xDescServ']),

                        # ---------------- VALORES ----------------
                        'Valor Serviço': pegar(dps, ['valores','vServPrest', 'vServ' ], 0),
                        'Descontos': pegar(valores, ['vDescCondIncond']),

                        #-----------------ISS----------------------
                        'Base ISS': pegar(valores, ['vBC'], 0),
                        'Aliq. ISS': pegar(valores, ['pAliqAplic'], 0),
                        'Valor ISSQN': pegar(valores, ['vISSQN']),

                        #---------------PIS e COFINS----------------

                        'CST PIS/COFINS': pegar(trib_fed, ['piscofins', 'CST']),
                        'Base PIS/COFINS': pegar(trib_fed, ['piscofins', 'vBCPisCofins']),
                        'Aliq.PIS': pegar(trib_fed, ['piscofins', 'pAliqPis']),
                        'Valor PIS': pegar(trib_fed, ['piscofins', 'vPis']),
                        'Aliq.COFINS': pegar(trib_fed, ['piscofins', 'pAliqCofins']),
                        'Valor COFINS': pegar(trib_fed, ['piscofins', 'vCofins']),
                        'Tipo de Retenção PC': pegar(trib_fed, ['piscofins', 'tpRetPisCofins']),

                        'IRRF': pegar(trib_fed, ['vRetIRRF']),
                        'CSLL': pegar(trib_fed, ['vRetCSLL']),

                        'Total retenções': pegar(valores, ['vTotalRet']),
                        'Valor Liquido': pegar(valores, ['vLiq']),

                        # UF
                        'Finalidade IBS e CBS': pegar(dps, ['IBSCBS', 'finNFSe']),
                        'Indicador de Operação': pegar(dps, ['IBSCBS', 'cIndOp']),
                        'cClasTrib': pegar(dps, ['IBSCBS', 'valores', 'trib', 'gIBSCBS' ,'cClassTrib'],'Não Destacado'),

                        'BC IBS e CBS': pegar(ibscbs, ['valores','vBC']),
                        'Aliquota IBS UF': pegar(ibscbs, ['valores', 'uf', 'pIBSUF']),
                        'Valor IBS UF': pegar(ibscbs, ['totCIBS', 'gIBS','gIBSUFTot', 'vIBSUF']),

                        'Aliquota IBS UF': pegar(ibscbs, ['valores', 'mun', 'pIBSMun']),
                        'Valor IBS UF': pegar(ibscbs, ['totCIBS', 'gIBS','gIBSMunTot', 'vIBSMun']),

                        'Aliquota IBS UF': pegar(ibscbs, ['valores', 'fed', 'pCBS']),
                        'Valor IBS UF': pegar(ibscbs, ['totCIBS', 'gCBS','vCBS']),


                    })


        except Exception as e:
            erros.append({'arquivo': arquivo, 'Erro': str(e), 'Detalhe': traceback.format_exc()})
            

    processados = pd.DataFrame(notas)
    servicos = pd.DataFrame(nfse)
    df_erros = pd.DataFrame(erros)
    agora = datetime.now().strftime("%d_%m_%Y %H_%M_%S")
    if not df_erros.empty:
        df_erros.to_csv(rf'{pasta_saida}\erros_{agora}.txt', index=False)
    if len(processados) > 0:
        processados['Emissão'] = pd.to_datetime(processados['Emissão'], utc=True).dt.strftime('%d/%m/%Y')
    if len(servicos) > 0:
        servicos['Emissão'] = pd.to_datetime(servicos['Emissão'], utc=True).dt.strftime('%d/%m/%Y')
        servicos['Competencia'] = pd.to_datetime(servicos['Competencia'], utc=True).dt.strftime('%d/%m/%Y')

    return {'nfe': processados, 
            'nfse' : servicos}

def gerar_unico_arquivo():
    '''
    Gera arquivo único caso não seja adicionado CNPJ
    '''
    global pasta_saida
    formato = tipo_saida.get()
    nome = nome_arquivo_entry.get()
    dados = leitor_xml()
    processados = dados['processados']
    nfse = dados['nfse']
    if formato == 'excel':
        processados.to_excel(rf'{pasta_saida}\{nome}.xlsx', index=False, header = True)
        nfse.to_excel(rf'{pasta_saida}\{nome}-nfse.xlsx', index=False, header = True)


    elif formato == 'csv':
        processados.to_csv(rf'{pasta_saida}\{nome}.csv', index=False, header = True, sep=';', encoding='UTF-8')
        nfse.to_csv(rf'{pasta_saida}\{nome}-nfse.csv', index=False, header = True, sep=';', encoding='UTF-8')

    else:
        processados.to_csv(rf'{pasta_saida}\{nome}.txt', index=False, header = True, sep=';', encoding='UTF-8')
        nfse.to_csv(rf'{pasta_saida}\{nome}- nfse.txt', index=False, header = True, sep=';', encoding='UTF-8')


def gerar_arquivo_final(df,nome,formato):
    '''
    Gera os arquivos em TXT ou CSV conforme paâmetros
    '''
    nome_primario = nome_arquivo_entry.get()
    if formato == 'csv':
        df.to_csv(rf'{pasta_saida}\{nome_primario}-{nome}.csv',index=False, header= True, sep=';')

    else:
        df.to_csv(rf'{pasta_saida}\{nome_primario}-{nome}.txt',index=False, header= True, sep=';')


def gerar_separado():
    ''' 
    Gera os arquivos em excel
    As abas a serem geradas dependem das varíaveis marcadas 
    
    Caso não seja selecionado nenhuma variável ou as variáveis selecionadas não forem alimetadas, 
    retorna erro'''

    global pasta_saida
    formato = tipo_saida.get()
    cnpj = cnpj_empresa.get()
    saida = nf_saida.get()
    entrada = nf_entrada.get()
    clientes = clie.get()
    fornecedores = forn.get()
    produtos_entradas = prod_com.get()
    produtos_saidas = prod_ven.get()
    cbenef_saida = cben_saida.get()
    formato = tipo_saida.get()
    ibscbs = ibscbs_check.get()
    analise_icms_saida_var = analise_icms_saida.get()
    analise_icms_entrada_var = analise_icms_entrada.get()
    analise_piscofins_var = analise_piscofins.get()
    analise_ipi_saida_var = analise_ipi_saida.get()
    analise_iss_saida_var = analise_iss_saida.get()

    nfs = nfs_check.get()
    leitura = leitor_xml()
    processados = leitura['nfe'] if not leitura['nfe'].empty else pd.DataFrame()
    nfse = leitura['nfse'] if not leitura['nfse'].empty else pd.DataFrame()
    nome_arquivo = nome_arquivo_entry.get()
    valida_multi_filiais = multi_filial.get()

    cnpj = cnpj.replace('/','').replace('.','').replace('-','')
    
    if valida_multi_filiais:
        cnpj = cnpj[:8]
        print(cnpj)

    if not nome_arquivo:
        nome_arquivo = 'XML Processados'

    if formato == 'excel':

        with pd.ExcelWriter(rf'{pasta_saida}\{nome_arquivo}.xlsx', engine='openpyxl') as writer:

            if not processados.empty:
                if saida:
                    nome = "Saidas"
                    df_saidas = processados.loc[processados['CNPJ Emitente'].str.startswith(cnpj)]
                    df_saidas.to_excel(writer, sheet_name = nome, index=False)

                if entrada:
                    nome = "Entradas"
                    df_entradas = processados.loc[processados['CNPJ Destinatário'].str.startswith(cnpj)]
                    df_entradas.to_excel(writer, sheet_name = nome, index=False)

                if fornecedores:
                    nome = 'Fornecedores'
                    entradas_forn = processados.loc[processados['CNPJ Destinatário'].str.startswith(cnpj)]
                    df_fornecedores_lista = entradas_forn[['CNPJ Emitente','Razão Social Emitente', 'IE Emitente','UF Emitente','Mun Emitente','Regime Trib']]
                    df_fornecedores_lista = df_fornecedores_lista.rename(columns={'CNPJ Emitente': 'CNPJ',
                                                                            'Razão Social Emitente': 'Razão Social',
                                                                            'IE Emitente': 'Inscrição Estadual',
                                                                            'UF Emitente': 'Estado',
                                                                            'Mun Emitente': 'Município',
                                                                            'Regime Trib': 'Regime Tributário',
                                                                            })

                    df_fornecedores_lista = df_fornecedores_lista.drop_duplicates(subset=['CNPJ'])
                    df_fornecedores_lista.to_excel(writer, sheet_name = nome, index=False)

                if clientes:
                    nome = 'Clientes'
                    entradas_clientes = processados.loc[processados['CNPJ Emitente'].str.startswith(cnpj)]


                    df_clientes_lista = entradas_clientes[['CNPJ Destinatário','Razão Social Destinatário', 'IE Destinatário','UF Destinatário','Mun Destinatário',]]
                    df_clientes_lista = df_clientes_lista.rename(columns={'CNPJ Destinatário': 'CNPJ',
                                                                            'Razão Social Destinatário': 'Razão Social',
                                                                            'IE Destinatário': 'Inscrição Estadual',
                                                                            'UF Destinatário': 'Estado',
                                                                            'Mun Destinatário': 'Município',
                                                                            })

                    df_clientes_lista = df_clientes_lista.drop_duplicates(subset=['CNPJ'])

                    df_clientes_lista.to_excel(writer, sheet_name = nome, index=False)

                if produtos_entradas:
                    nome = 'Produtos Entradas'
                    todas_entradas = processados.loc[processados['CNPJ Destinatário'].str.startswith(cnpj)]


                    df_produtos_entradas = todas_entradas[['Cod.Produto', 'Descrição', 'NCM', 'UM', 'Origem']]
                    df_produtos_entradas = df_produtos_entradas.drop_duplicates(subset=['Cod.Produto'])

                    df_produtos_entradas.to_excel(writer, sheet_name = nome, index=False)

                if produtos_saidas: 
                    nome = 'Produtos Saidas'
                    todas_saidas = processados.loc[processados['CNPJ Emitente'].str.startswith(cnpj)]


                    df_produtos_saidas = todas_saidas[['Cod.Produto', 'Descrição', 'NCM', 'UM', 'Origem']]
                    df_produtos_saidas = df_produtos_saidas.drop_duplicates(subset=['Cod.Produto'])

                    df_produtos_saidas.to_excel(writer, sheet_name = nome, index=False)

                if cbenef_saida:
                    nome = 'cBenef'
                    todas_saidas = processados.loc[processados['CNPJ Emitente'].str.startswith(cnpj)]


                    df_cbenef = todas_saidas[['Cod.Produto', 'Descrição', 'NCM', 'Origem', 'cBenef', 'CFOP', 'CST ICMS']]
                    df_cbenef = df_cbenef.drop_duplicates(subset=['Cod.Produto','cBenef', 'CFOP', 'CST ICMS'])
                    df_cbenef = df_cbenef.loc[df_cbenef['cBenef'].astype(str) != '']
                    df_cbenef.to_excel(writer, sheet_name = nome, index=False)

                if ibscbs:

                    nome = 'IBS e CBS'
                    todos_movimentos = processados
                    todos_movimentos['Operação'] = ''
                    todos_movimentos.loc[todos_movimentos['CNPJ Destinatário'].str.startswith(cnpj), 'Operação'] = 'Entrada'
                    todos_movimentos.loc[todos_movimentos['CNPJ Emitente'].str.startswith(cnpj), 'Operação'] = 'Saídas'


                    df_cbseibs = todos_movimentos[['Operação','Cod.Produto', 'Descrição', 'NCM', 'Origem', 'CFOP','CCLASTRIB','Aliq. IBSUF','Aliq. IBSMun', 'Aliq. CBS']]
                    df_cbseibs = df_cbseibs.drop_duplicates(subset=['Cod.Produto','CCLASTRIB', 'CFOP'])
                    df_cbseibs = df_cbseibs.loc[df_cbseibs['CCLASTRIB'].astype(str) != '']

                    df_cbseibs.to_excel(writer, sheet_name = nome, index=False)

                if analise_icms_saida_var:
                    nome = 'Analise ICMS Saída'
                    df_saidas_icms = processados.loc[processados['CNPJ Emitente'].str.startswith(cnpj)]

                    resumo_icms_saida = df_saidas_icms[['CFOP', 'CST ICMS','UF Emitente', 'UF Destinatário', '% ICMS', '% Redução BC', '% Diferimento ICMS', 'MVA ICMS ST', 'cBenef']]
                    resumo_icms_saida = resumo_icms_saida.groupby(['CFOP', 'CST ICMS','UF Emitente', 'UF Destinatário', '% ICMS', '% Redução BC', '% Diferimento ICMS', 'MVA ICMS ST', 'cBenef']).size().reset_index(name='Contagem')
                    resumo_icms_saida = resumo_icms_saida.sort_values('Contagem', ascending=False)
                    resumo_icms_saida.to_excel(writer, sheet_name = nome, index=False)

                if analise_icms_entrada_var:
                    nome = 'Analise ICMS Entradas'
                    df_entradas_icms = processados.loc[processados['CNPJ Destinatário'].str.startswith(cnpj)]

                    resumo_icms_entrada = df_entradas_icms[['CFOP', 'CST ICMS','UF Emitente', 'UF Destinatário', '% ICMS', '% Redução BC', '% Diferimento ICMS', 'MVA ICMS ST', 'cBenef']]
                    resumo_icms_entrada = resumo_icms_entrada.groupby(['CFOP', 'CST ICMS','UF Emitente', 'UF Destinatário', '% ICMS', '% Redução BC', '% Diferimento ICMS', 'MVA ICMS ST', 'cBenef']).size().reset_index(name='Contagem')
                    resumo_icms_entrada = resumo_icms_entrada.sort_values('Contagem', ascending=False)
                    resumo_icms_entrada.to_excel(writer, sheet_name = nome, index=False)
                
                if analise_piscofins_var:
                    print(analise_piscofins_var)
                    nome = 'Analise PIS e COFINS Saída'
                    dfs = []
                    if not processados.empty:
                        df_saidas_piscofins = processados.loc[processados['CNPJ Emitente'].str.startswith(cnpj)]
                        resumo_pc_saida_nf = df_saidas_piscofins[['CFOP','UF Emitente', 'CST PIS','% PIS', 'CST COFINS', '% COFINS']]
                        dfs.append(resumo_pc_saida_nf)
                        

                    if not nfse.empty:
                        df_saidas_nfs_pc = nfse.loc[nfse['CNPJ Emitente'].str.startswith(cnpj)]
                        resumo_pc_saida_nfs = df_saidas_nfs_pc[['CST PIS/COFINS', 'UF Prestador', 'Aliq.PIS', 'Aliq.COFINS', 'Tipo de Retenção PC']]
                        resumo_pc_saida_nfs = resumo_pc_saida_nfs.rename(columns={'CST PIS/COFINS': 'CST PIS', 
                                                                                  'UF Prestador': 'UF Emitente',
                                                                                  'Aliq.PIS': '% PIS', 
                                                                                  'Aliq.COFINS': '% COFINS'})
                        resumo_pc_saida_nfs['CFOP'] = '5933/6933'
                        resumo_pc_saida_nfs['CST COFINS'] = resumo_pc_saida_nfs['CST PIS']
                        dfs.append(resumo_pc_saida_nfs)

                    if dfs:
                        colunas = ['CFOP','UF Emitente','CST PIS','% PIS', 'CST COFINS', '% COFINS', 'Tipo de Retenção PC']
                        resumo_pc_saida = pd.concat(dfs, ignore_index=True)
                        resumo_pc_saida = resumo_pc_saida.reindex(columns=colunas)
                        resumo_pc_saida = resumo_pc_saida.groupby(['CFOP','UF Emitente','CST PIS','% PIS', 'CST COFINS', '% COFINS', 'Tipo de Retenção PC'], dropna=False).size().reset_index(name='Contagem')
                        resumo_pc_saida = resumo_pc_saida.sort_values('Contagem', ascending=False)

                        resumo_pc_saida.to_excel(writer, sheet_name = nome, index=False)

                if analise_ipi_saida_var:
                    nome = 'Analise IPI'
                    movimentos = processados
                    movimentos['Operação'] = '' 
                    movimentos.loc[movimentos['CNPJ Destinatário'].str.startswith(cnpj), 'Operação'] = 'Entrada'
                    movimentos.loc[movimentos['CNPJ Emitente'].str.startswith(cnpj), 'Operação'] = 'Saídas'
                    resumo_ipi = movimentos[['Operação','CFOP', 'NCM', 'Enq.IPI', 'CST IPI', '% IPI']]
                    resumo_ipi = resumo_ipi.groupby(['Operação','CFOP', 'NCM', 'Enq.IPI', 'CST IPI', '% IPI'], dropna=False).size().reset_index(name='Contagem')
                    resumo_ipi = resumo_ipi.sort_values(by=['Operação', 'Contagem'], ascending=False)

                    resumo_ipi.to_excel(writer, sheet_name = nome, index=False)




            if not nfse.empty:
                if nfs:
                    nome = 'NFSe emitidas'
                    df_saidas_nfs = nfse.loc[nfse['CNPJ Emitente'].str.startswith(cnpj)]
                    if not df_saidas_nfs.empty:
                        df_saidas_nfs.to_excel(writer, sheet_name = nome, index=False)

                    nome = 'NFSe recebidas'
                    df_entrada_nfs = nfse.loc[nfse['CNPJ Tomador'].str.startswith(cnpj)]
                    if not df_entrada_nfs.empty:
                        df_entrada_nfs.to_excel(writer, sheet_name = nome, index=False)

                if analise_iss_saida_var:
                    nome = 'Analise ISS'
                    movimentos = nfse
                    movimentos['Operação'] = ''
                    movimentos.loc[movimentos['CNPJ Tomador'].str.startswith(cnpj), 'Operação'] = 'Entrada'
                    movimentos.loc[movimentos['CNPJ Emitente'].str.startswith(cnpj), 'Operação'] = 'Saídas'
                    resumo_iss = movimentos[['Operação', 'Local Prestação', 'Código Serviço', 'Código NBS', 'Cod.Trib. Municipal', 'Aliq. ISS']]
                    resumo_iss = resumo_iss.groupby(['Operação', 'Local Prestação', 'Código Serviço', 'Código NBS', 'Cod.Trib. Municipal', 'Aliq. ISS'], dropna=False).size().reset_index(name='Contagem')
                    resumo_iss = resumo_iss.sort_values(by=['Operação', 'Contagem'], ascending=False)
                    if not resumo_iss.empty:
                        resumo_iss.to_excel(writer, sheet_name = nome, index=False)
                        nfse.to_excel(writer, sheet_name = 'todas', index=False)


    
    else:
        '''
        Tratativa para vários arquivos CSV ou TXT
        '''
        if not processados.empty:
            if saida:
                nome = "Saidas"
                df_saidas = processados.loc[processados['CNPJ Emitente'].str.startswith(cnpj)]
                gerar_arquivo_final(df_saidas, nome, formato)

            if entrada:
                nome = "Entradas"
                df_entradas = processados.loc[processados['CNPJ Destinatário'].str.startswith(cnpj)]
                gerar_arquivo_final(df_entradas, nome, formato)

            if fornecedores:
                nome = 'Fornecedores'
                entradas_forn = processados.loc[processados['CNPJ Destinatário'].str.startswith(cnpj)]
                df_fornecedores_lista = entradas_forn[['CNPJ Emitente','Razão Social Emitente', 'IE Emitente','UF Emitente','Mun Emitente','Regime Trib']]
                df_fornecedores_lista = df_fornecedores_lista.rename(columns={'CNPJ Emitente': 'CNPJ',
                                                                        'Razão Social Emitente': 'Razão Social',
                                                                        'IE Emitente': 'Inscrição Estadual',
                                                                        'UF Emitente': 'Estado',
                                                                        'Mun Emitente': 'Município',
                                                                        'Regime Trib': 'Regime Tributário',
                                                                        })
                df_fornecedores_lista = df_fornecedores_lista.drop_duplicates(subset=['CNPJ'])
                gerar_arquivo_final(saida, nome, formato)

            if clientes:
                nome = 'Clientes'
                entradas_clientes = processados.loc[processados['CNPJ Emitente'].str.startswith(cnpj)]
                df_clientes_lista = entradas_clientes[['CNPJ Destinatário','Razão Social Destinatário', 'IE Destinatário','UF Destinatário','Mun Destinatário',]]
                df_clientes_lista = df_clientes_lista.rename(columns={'CNPJ Destinatário': 'CNPJ',
                                                                        'Razão Social Destinatário': 'Razão Social',
                                                                        'IE Destinatário': 'Inscrição Estadual',
                                                                        'UF Destinatário': 'Estado',
                                                                        'Mun Destinatário': 'Município',
                                                                        })
                df_clientes_lista = df_clientes_lista.drop_duplicates(subset=['CNPJ'])
                gerar_arquivo_final(df_clientes_lista, nome, formato)

            if produtos_entradas:
                nome = 'Produtos Entradas'
                todas_entradas = processados.loc[processados['CNPJ Destinatário'].str.startswith(cnpj)]
                df_produtos_entradas = todas_entradas[['Cod.Produto', 'Descrição', 'NCM', 'UM', 'Origem']]
                df_produtos_entradas = produtos_entradas.drop_duplicates(subset=['Cod.Produto'])
                gerar_arquivo_final(df_produtos_entradas, nome, formato)

            if produtos_saidas: 
                nome = 'Produtos Saidas'
                todas_saidas = processados.loc[processados['CNPJ Emitente'].str.startswith(cnpj)]
                df_produtos_saidas = todas_saidas[['Cod.Produto', 'Descrição', 'NCM', 'UM', 'Origem']]
                df_produtos_saidas = produtos_saidas.drop_duplicates(subset=['Cod.Produto'])
                gerar_arquivo_final(df_produtos_saidas, nome, formato)

            if cbenef_saida:
                nome = 'cBenef'
                todas_saidas = processados.loc[processados['CNPJ Emitente'].str.startswith(cnpj)]
                df_cbenef = todas_saidas[['Cod.Produto', 'Descrição', 'NCM', 'Origem', 'cBenef', 'CFOP', 'CST ICMS']]
                df_cbenef = df_cbenef.drop_duplicates(subset=['Cod.Produto','cBenef', 'CFOP', 'CST ICMS'])
                df_cbenef = df_cbenef.loc[df_cbenef['cBenef'].astype(str) != '']
                gerar_arquivo_final(df_cbenef, nome, formato)

            if ibscbs_check:
                nome = 'IBS e CBS'
                todos_movimentos = processados
                todos_movimentos['Operação'] = ''
                todos_movimentos.loc[todos_movimentos['CNPJ Destinatário'].str.startswith(cnpj), 'Operação'] = 'Entrada'
                todos_movimentos.loc[todos_movimentos['CNPJ Emitente'].str.startswith(cnpj), 'Operação'] = 'Saídas'
                df_cbseibs = todos_movimentos[['Operação','Cod.Produto', 'Descrição', 'NCM', 'Origem', 'CFOP','CCLASTRIB','Aliq. IBSUF','Aliq. IBSMun', 'Aliq. CBS']]
                df_cbseibs = df_cbseibs.drop_duplicates(subset=['Cod.Produto','CCLASTRIB', 'CFOP'])
                df_cbseibs = df_cbseibs.loc[df_cbseibs['CCLASTRIB'].astype(str) != '']
                gerar_arquivo_final(df_cbseibs, nome, formato)

                if analise_icms_saida_var:
                    nome = 'Analise ICMS Saída'
                    df_saidas_icms = processados.loc[processados['CNPJ Emitente'].str.startswith(cnpj)]

                    resumo_icms_saida = df_saidas_icms[['CFOP', 'CST ICMS','UF Emitente', 'UF Destinatário', '% ICMS', '% Redução BC', '% Diferimento ICMS', 'MVA ICMS ST', 'cBenef']]
                    resumo_icms_saida = resumo_icms_saida.groupby(['CFOP', 'CST ICMS','UF Emitente', 'UF Destinatário', '% ICMS', '% Redução BC', '% Diferimento ICMS', 'MVA ICMS ST', 'cBenef']).size().reset_index(name='Contagem')
                    resumo_icms_saida = resumo_icms_saida.sort_values('Contagem', ascending=False)
                    gerar_arquivo_final(resumo_icms_saida, nome, formato)

                if analise_icms_entrada_var:
                    nome = 'Analise ICMS Entradas'
                    df_entradas_icms = processados.loc[processados['CNPJ Destinatário'].str.startswith(cnpj)]

                    resumo_icms_entrada = df_entradas_icms[['CFOP', 'CST ICMS','UF Emitente', 'UF Destinatário', '% ICMS', '% Redução BC', '% Diferimento ICMS', 'MVA ICMS ST', 'cBenef']]
                    resumo_icms_entrada = resumo_icms_entrada.groupby(['CFOP', 'CST ICMS','UF Emitente', 'UF Destinatário', '% ICMS', '% Redução BC', '% Diferimento ICMS', 'MVA ICMS ST', 'cBenef']).size().reset_index(name='Contagem')
                    resumo_icms_entrada = resumo_icms_entrada.sort_values('Contagem', ascending=False)
                    gerar_arquivo_final(resumo_icms_saida, nome, formato)

                if analise_piscofins_var:
                    nome = 'Analise PIS e COFINS Saída'
                    if not processados.empty:
                        df_saidas_piscofins = processados.loc[processados['CNPJ Emitente'].str.startswith(cnpj)]
                        resumo_pc_saida_nf = df_saidas_piscofins[['CFOP','UF Emitente', 'CST PIS','% PIS', 'CST COFINS', '% COFINS']]

                    if not nfse.empty:
                        df_saidas_nfs_pc = nfse.loc[nfse['CNPJ Emitente'].str.startswith(cnpj)]
                        resumo_pc_saida_nfs = df_saidas_nfs_pc[['CST PIS/COFINS', 'UF Prestador', 'Aliq.PIS', 'Aliq.COFINS', 'Tipo de Retenção PC']]
                        resumo_pc_saida_nfs = resumo_pc_saida_nfs.rename(columns={'CST PIS/COFINS': 'CST PIS', 
                                                                                  'UF Prestador': 'UF Emitente',
                                                                                  'Aliq.PIS': '% PIS', 
                                                                                  'Aliq.COFINS': '% COFINS'})
                        resumo_pc_saida_nfs['CFOP'] = '5933/6933'
                        resumo_pc_saida_nfs['CST COFINS'] = resumo_pc_saida_nfs['CST PIS']

                    if not resumo_pc_saida_nfs.empty and not resumo_pc_saida_nf.empty:
                        resumo_pc_saida = pd.concat([resumo_pc_saida_nf, resumo_pc_saida_nfs], ignore_index=True)

                    elif not resumo_pc_saida_nfs.empty and resumo_pc_saida_nf.empty:
                        resumo_pc_saida = resumo_pc_saida_nfs
                    elif resumo_pc_saida_nfs.empty and  not resumo_pc_saida_nf.empty:
                        resumo_pc_saida = resumo_pc_saida_nf

                    if not resumo_pc_saida.empty:
                        resumo_pc_saida = resumo_pc_saida.groupby(['CFOP','UF Emitente','CST PIS','% PIS', 'CST COFINS', '% COFINS', 'Tipo de Retenção PC'], dropna=False).size().reset_index(name='Contagem')
                        resumo_pc_saida = resumo_pc_saida.sort_values('Contagem', ascending=False)

                        gerar_arquivo_final(resumo_pc_saida, nome, formato)

                    if analise_ipi_saida_var:
                        nome = 'Analise IPI'
                        movimentos = processados
                        movimentos['Operação'] = '' 
                        movimentos.loc[movimentos['CNPJ Destinatário'].str.startswith(cnpj), 'Operação'] = 'Entrada'
                        movimentos.loc[movimentos['CNPJ Emitente'].str.startswith(cnpj), 'Operação'] = 'Saídas'
                        resumo_ipi = movimentos[['Operação','CFOP', 'NCM', 'Enq.IPI', 'CST IPI', '% IPI']]
                        resumo_ipi = resumo_ipi.groupby(['Operação','CFOP', 'NCM', 'Enq.IPI', 'CST IPI', '% IPI'], dropna=False).size().reset_index(name='Contagem')
                        resumo_ipi = resumo_ipi.sort_values(by=['Operação', 'Contagem'], ascending=False)
                        gerar_arquivo_final(resumo_ipi, nome, formato)

        if not nfse.empty:
            if nfs:
                nome = 'NFSe emitidas'
                df_saidas_nfs = nfse.loc[nfse['CNPJ Emitente'].str.startswith(cnpj)]
                df_saidas_nfs.to_excel(writer, sheet_name = nome, index=False)
                gerar_arquivo_final(df_saidas_nfs, nome, formato)

                nome = 'NFSe recebidas'
                df_entrada_nfs = nfse.loc[nfse['CNPJ Tomador'].str.startswith(cnpj)]
                df_entrada_nfs.to_excel(writer, sheet_name = nome, index=False)
                gerar_arquivo_final(df_saidas_nfs, nome, formato)

                if analise_iss_saida_var:
                    nome = 'Analise ISS'
                    movimentos = nfse
                    movimentos['Operação'] = ''
                    movimentos.loc[movimentos['CNPJ Tomador'].str.startswith(cnpj), 'Operação'] = 'Entrada'
                    movimentos.loc[movimentos['CNPJ Emitente'].str.startswith(cnpj), 'Operação'] = 'Saídas'
                    resumo_iss = movimentos[['Operação', 'Local Prestação', 'Código Serviço', 'Código NBS', 'Cod.Trib. Municipal', 'Aliq. ISS']]
                    resumo_iss = resumo_iss.groupby(['Operação', 'Local Prestação', 'Código Serviço', 'Código NBS', 'Cod.Trib. Municipal', 'Aliq. ISS'], dropna=False).size().reset_index(name='Contagem')
                    resumo_iss = resumo_iss.sort_values(by=['Operação', 'Contagem'], ascending=False)
                    if not resumo_iss.empty:
                        gerar_arquivo_final(resumo_iss, nome, formato)

def escolher_pasta():
    global pasta_xml
    global quantidade
    pasta_xml = filedialog.askdirectory(
        title="Origem dos XMLs:"
        )

    if pasta_xml:
        pasta = Path(pasta_xml)
        quantidade = len(list(pasta.rglob("*.xml")))
        
        return frame1_origem.config(text=f'Pasta Selecionada: {pasta} \n Quantidade de XMLs na pasta: {quantidade}')
    
def escolher_pasta_salvar():
    global pasta_saida
    pasta = filedialog.askdirectory(
        title="Salvar em:"
        )

    if pasta:
        pasta_saida = pasta
        return frame1_destino.config(text=pasta)
    
# def verifica_excel():
#     if tipo_saida.get() == 'excel':
#         check_excel.config(state='normal')
#     else:
#         check_excel.config(state='disabled')

def gerar_arquvio(tipo):
    try:
        if tipo == 'unico':
            gerar_unico_arquivo()
        else:
            gerar_separado()
        root.after(0, finalizar)
    except Exception as e:
        erro = traceback.format_exc()
        root.after(0, lambda: finalizar_erro(erro))


def mascara_cnpj(event):
    texto = cnpj_empresa.get()
    numeros = ''.join(filter(str.isdigit, texto))[:14]

    novo = ""

    if len(numeros) >= 1:
        novo += numeros[:2]
    if len(numeros) >= 3:
        novo += "." + numeros[2:5]
    if len(numeros) >= 6:
        novo += "." + numeros[5:8]
    if len(numeros) >= 9:
        novo += "/" + numeros[8:12]
    if len(numeros) >= 13:
        novo += "-" + numeros[12:14]

    cnpj_empresa.delete(0, "end")
    cnpj_empresa.insert(0, novo)

def iniciar(tipo):
    #Tela de Load
    global loading, progress
    loading = tb.Toplevel(root)
    loading.title("Processando...")
    loading.geometry("350x150")
    loading.resizable(False,False)

    tb.Label(loading, text="Processando XML...").pack(pady=10)

    progress = tb.Progressbar(loading, mode="indeterminate", bootstyle=SUCCESS)
    progress.pack(fill="x", padx=20, pady=10)

    progress.start()

    #Iniciar Tread
    thread = threading.Thread(target=gerar_arquvio, args=(tipo,))
    thread.start()

def finalizar():
    progress.stop()
    loading.destroy()
    Messagebox.show_info('Arquivo Gerado com Sucesso', 'Sucesso!')

def finalizar_erro(erro):
    progress.stop()
    loading.destroy()
    erro_win = tb.Toplevel()
    erro_win.title("Erro Crítico")
    erro_win.geometry("500x300")  # controla tamanho

    tb.Label(erro_win, text="Erro critico ao processar:", bootstyle=DANGER).pack(pady=10)

    txt = tb.Text(erro_win, height=10)
    txt.pack(fill=BOTH, expand=True, padx=10, pady=10)

    txt.insert("1.0", erro)

    tb.Button(erro_win, text="Fechar", command=erro_win.destroy).pack(pady=10)

    # data = datetime.now().strftime("%d_%m_%Y %H_%M_%S")
    # with open (f'erro_processo_{data}.txt', 'w', encoding='utf-8') as e:
    #     e.write(erro)

def nome_valido(nome):
    invalidos = r'<>:"/\|?*'

    if not nome:
        return False
    
    if any(c in nome for c in invalidos):
        return False
    
    if nome.endswith(" ") or nome.endswith("."):
        return False
    
    reservados = {
        "CON","PRN","AUX","NUL",
        "COM1","COM2","COM3","COM4","COM5","COM6","COM7","COM8","COM9",
        "LPT1","LPT2","LPT3","LPT4","LPT5","LPT6","LPT7","LPT8","LPT9"
    }
    
    if nome.upper() in reservados:
        return False
    
    return True

def validacao():
    cnpj = cnpj_empresa.get()
    formato = tipo_saida.get()
    nome = nome_arquivo_entry.get()
    caracteres = nome_valido(nome)
    if caracteres:
        if pasta_xml:
            if pasta_saida:
                if quantidade > 0:
                    if cnpj:
                        iniciar('separado')
                    else:
                        resposta = Messagebox.show_question('O Campo CNPJ não foi preenchido. \n Será gerado uma unica tabela com todos os documentos da pasta, sem separações \n Deseja Continuar?', 'Confirmação', buttons=['Não:secundary', 'Sim:primary'])
                        
                        if resposta == 'Não':
                            pass
                        else: 
                            iniciar('unico')
                else:
                    Messagebox.show_error('A Pasta selecionada não contém nenhum XML, favor revisar!','Erro')
                    
            else:
                Messagebox.show_error('Não foi Selecionada pasta Destino dos arquivos','Erro')
                
        else:
            Messagebox.show_error('Não foi Selecionada pasta com os XMLs','Erro')
    else:
        Messagebox.show_error('Nome do arquivo Vazio ou inválido \n O nome do arquvo não pode estar vazio ou conter os caracteres:' + r'<>:"/\|?*' +'\n Favor corrigir antes de seguir.','Erro')
            
def inverter():
    if marcar_desmarcar.get():
        for var in var_check:
            var.set(True)
    else:
        for var in var_check:
            var.set(False)

# Janela TTK:

#Variáveis: 



root = tb.Window(themename='solar')
root.title("Análise XML - Relatório Fiscal. v1.0")
root.geometry('800x950')

nf_saida = tb.BooleanVar(value=True)
nf_entrada = tb.BooleanVar(value=True)
forn = tb.BooleanVar(value=True)
clie = tb.BooleanVar(value=True)
prod_com = tb.BooleanVar(value=True)
prod_ven = tb.BooleanVar(value=True)
cben_saida = tb.BooleanVar(value=True)
cben_entradas = tb.BooleanVar(value=True)
ibscbs_check = tb.BooleanVar(value=True)
tipo_saida = tb.StringVar(value='excel')
#unico_excel = tb.BooleanVar()
multi_filial = tb.BooleanVar()
nfs_check = tb.BooleanVar(value=True)
analise_icms_saida = tb.BooleanVar(value=True)
analise_icms_entrada = tb.BooleanVar(value=True)
analise_piscofins = tb.BooleanVar(value=True)
analise_ipi_saida = tb.BooleanVar(value=True)
analise_iss_saida = tb.BooleanVar(value=True)
marcar_desmarcar = tb.BooleanVar(value=True)

var_check = [nfs_check, nf_saida, nf_entrada, forn, clie, prod_com, prod_ven, cben_saida, ibscbs_check, analise_icms_saida, analise_icms_entrada, analise_ipi_saida, analise_iss_saida, analise_piscofins]


frame1 = tb.Frame(root)
frame1.pack(fill=BOTH, expand=True)

frame1_inicio = tb.Label(frame1, justify='center', anchor='center', text='Aplicativo para ler XMLs de Compras e Vendas \n e gerar um relatório em Excel, CSV ou TXT para \n facilitar no levantamento fiscal', bootstyle=INFO)
frame1_inicio.pack(pady=10, fill=X) 

frame1_sep1 = tb.Separator(frame1, orient="horizontal")
frame1_sep1.pack(fill=X, pady=5)

#Dados da Empresa:
frame1_empresa = tb.Frame(frame1)
frame1_empresa.pack(pady=10)

frame1_empresa_txt = tb.Label(frame1_empresa, justify='center', anchor='center', text='Digite o CNPJ da Empresa para separar as operações em Vendas e Compras \n Permitindo importar todos os XMLs de uma Unica Vez (Compras e Vendas)', bootstyle=INFO)
frame1_empresa_txt.grid(row=0, column=0, padx=20, pady=10, columnspan=2)

cnpj_empresa = tb.Entry(frame1_empresa, text='CNPJ')
cnpj_empresa.grid(row=1, column=0, padx=20, pady=10)
cnpj_empresa.bind("<KeyRelease>", mascara_cnpj)

check_empresa = tb.Checkbutton(frame1_empresa, text='Considera Multi-Filiais?', variable=multi_filial)
check_empresa.grid(row=1, column=1, padx=10, pady=10)

ToolTip(check_empresa, text="""
        Ao considerar Multi Filiais a validação do CNPJ ocorrerá pelos oito primeiros digitos,.
        Caso essa opção estiver desmarcada irá validar o CNPJ completo.
        """)

frame1_empresa_txt2 = tb.Label(frame1_empresa, justify='center', anchor='center', text='', bootstyle=INFO)
frame1_empresa_txt2.grid(row=0, column=0, padx=20, pady=10, columnspan=2)

frame1_sep2 = tb.Separator(frame1, orient="horizontal")
frame1_sep2.pack(fill=X, pady=5)


frame1_button = tb.Button(frame1, text="Selecionar pasta XML:", command=escolher_pasta, bootstyle=PRIMARY)
frame1_button.pack(pady=10, padx=20, fill=X)



frame1_origem = tb.Label(frame1, text="Pasta não selecionada!", bootstyle=INFO )
frame1_origem.pack(pady=10)

frame1_sep3 = tb.Separator(frame1, orient="horizontal")
frame1_sep3.pack(fill=X, pady=5)

frame1_arquivo = tb.Label(frame1, text="Digite o nome do arquivo a ser salvo, sem a extensão:", bootstyle=INFO )
frame1_arquivo.pack(pady=10)

nome_arquivo_entry = tb.Entry(frame1, text='nome do Arquivo:')
nome_arquivo_entry.pack(pady=10, padx=20, fill='x')

frame1_sep4 = tb.Separator(frame1, orient="horizontal")
frame1_sep4.pack(fill=X, pady=5)
#config: 
frame_check = tb.Frame(frame1)
frame_check.pack(pady=10)

frame_check_txt = tb.Label(frame_check, text='informações a serem geradas:')
frame_check_txt.grid(row=0, column=0, padx=20, pady=10, columnspan=4)

tb.Checkbutton(frame_check, text="Tabela de Saídas", variable=nf_saida).grid(row=1, column=0,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Tabela de Entradas", variable=nf_entrada).grid(row=1, column=1,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Lista de Fornecedores", variable=forn).grid(row=1, column=2,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Lista de Clientes", variable=clie).grid(row=1, column=3,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Produtos de Compras", variable=prod_com).grid(row=2, column=0,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Produtos de Vendas", variable=prod_ven).grid(row=2, column=1,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="CBENEF nas Vendas", variable=cben_saida).grid(row=2, column=2,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Regras de CBS e IBS NFe", variable=ibscbs_check).grid(row=2, column=3,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Notas de Serviço (NFSe)", variable=nfs_check).grid(row=3, column=0,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Análise ICMS Saídas", variable=analise_icms_saida).grid(row=3, column=1,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Análise ICMS Entradas", variable=analise_icms_entrada).grid(row=3, column=2,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Análise PIS/COFINS Saídas", variable=analise_piscofins).grid(row=3, column=3,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Análise IPI", variable=analise_ipi_saida).grid(row=4, column=0,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Análise ISS", variable=analise_iss_saida).grid(row=4, column=1,sticky='w', padx=10, pady=5)
tb.Checkbutton(frame_check, text="Marcar/Desmarcar tudo", variable=marcar_desmarcar, command=inverter).grid(row=6, column=0,sticky='w', padx=10, pady=5)

frame1_sep4 = tb.Separator(frame1, orient="horizontal")
frame1_sep4.pack(fill=X, pady=5)

#Tipo de arquivo

frame_check2 = tb.Frame(frame1)
frame_check2.pack(pady=10)

frame_check2_txt = tb.Label(frame_check2, text='Tipo de saída do Arquivo:')
frame_check2_txt.grid(row=0, column=0, padx=20, pady=10, columnspan=3)

tb.Radiobutton(frame_check2, text="Excel", value="excel", variable=tipo_saida).grid(row=1, column=0,sticky='w', padx=20, pady=5)
tb.Radiobutton(frame_check2, text="CSV", value="csv", variable=tipo_saida).grid(row=1, column=1,sticky='w', padx=20, pady=5)
tb.Radiobutton(frame_check2, text="TXT", value="txt", variable=tipo_saida).grid(row=1, column=2,sticky='w', padx=20, pady=5)
#check_excel = tb.Checkbutton(frame_check2, text="Saida de um arquvo XLSX por tabela?", variable=unico_excel)
#check_excel.grid(row=2, column=0,sticky='w', padx=20, pady=5, columnspan=2)
#check_excel.config(state='disabled')

#pegar as variaveis: 
#print(tipo_saida.get())

frame1_sep5 = tb.Separator(frame1, orient="horizontal")
frame1_sep5.pack(fill=X, pady=5)

frame1_button2 = tb.Button(frame1, text="Salvar em:", command=escolher_pasta_salvar, bootstyle=PRIMARY)
frame1_button2.pack(pady=10, padx=20, fill=X)

frame1_destino = tb.Label(frame1, text="Pasta não selecionada!", bootstyle=INFO )
frame1_destino.pack(pady=10)

frame1_button3 = tb.Button(frame1, text="Gerar arquivos:", command=validacao, bootstyle=PRIMARY)
frame1_button3.pack(pady=10, padx=20, fill=X)

root.mainloop()