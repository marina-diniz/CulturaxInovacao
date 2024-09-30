# Automatizazão x Cultura

# Instalar pacotes necessários
install.packages(c("readxl", "dplyr", "rpart", "rpart.plot", "caTools", "e1071", "corrplot", "caret", "ggplot2","smotefamily", "psych","xpsych","openxlsx"))

# Carregar pacotes
library(readxl)
library(dplyr)
library(rpart)
library(rpart.plot)
library(caTools)
library(e1071)
library(corrplot)
library(ggplot2)
library(psych)
library(openxlsx)



# _______________________________________ Extração de dados da base e Renomear colunas 
# Carregar os dados do arquivo Excel
dados <- read_excel("Questionário de Cultura x Inovação (respostas) (3).xlsx")

# Renomear colunas
colnames(dados) <- c(
  "Data_hora",
  "Funcoes_administrativas",
  "Area_atuacao",
  "Nivel_hierarquico",
  "Setor_empresa",
  "Empresa_Publica_Privada",
  "Sexo",
  "Faixa_etaria",
  "Faixa_salarial",
  "Autonomia_iniciativa_tomada_de_decisoes",
  "Processos_claros",
  "Ambiente_de_colaboracao",
  "Ambiente_seguro_comunicacao_clara",
  "Estilo_de_lideranca",
  "Lideranca_exemplo_inovacao",
  "Ferramentas_para_automatizacoes",
  "Treinamento_capacitacao",
  "Investimento_tecnologia",
  "Qualidade_da_infraestrutura",
  "Plataformas_intuitivas",
  "Recompensa_iniciativas_automatizacao",
  "Empresa_agil",
  "Empresa_flexivel",
  "Cultura_aprendizado_continuo",
  "Trabalhos_repetitivos",
  "capaz_implementar_planilhas",
  "Implementar_planilhas_automatizadas",
  "apresentacao_de_dados",
  "analisa_dados_da_melhor_forma",
  "ferramentas_favorecem_a_inovacao",
  "ferramentas_voce_aplica"
)


# _______________________________________ Converção de Variáveis em Variáveis Numericas 

# Definir variáveis a serem convertidas
variaveis_convertidas <- c(
  "Autonomia_iniciativa_tomada_de_decisoes",
  "Processos_claros",
  "Ambiente_de_colaboracao",
  "Ambiente_seguro_comunicacao_clara",
  "Lideranca_exemplo_inovacao",
  "Treinamento_capacitacao",
  "Plataformas_intuitivas",
  "Recompensa_iniciativas_automatizacao",
  "Empresa_agil",
  "Empresa_flexivel",
  "Cultura_aprendizado_continuo",
  "Trabalhos_repetitivos"
)

# Função para converter variáveis categóricas para numéricas
converter_para_numerico <- function(coluna) {
  categorias <- c("Nunca", "Raramente", "Às vezes", "Frequentemente", "Sempre", "Às Vezes", "Nunca.","Não realizo trabalhos repetitivos")
  valores <- c(1, 2, 3, 4, 5, 3, 1, 0)
  as.numeric(factor(coluna, levels = categorias, labels = valores))
}

# Aplicar a conversão para as variáveis especificadas em dados
dados_tratados <- dados %>%
  mutate_at(vars(one_of(variaveis_convertidas)), converter_para_numerico)

# Definir vetor de correspondência para os novos rótulos de Estilo_de_lideranca
correspondencia_lideranca <- c(
  "Autoritário(a): concentra decisões e poder em si mesmo(a), com pouca autonomia para os colaboradores." = "Autoritario",
  "Democrático(a): busca consenso e envolve os colaboradores na tomada de decisões." = "Democratico",
  "Laissez-faire: concede ampla autonomia aos colaboradores, com pouca intervenção da liderança." = "Laissez-faire",
  "Transformacional(ista): inspira e motiva os colaboradores a alcançar objetivos desafiadores." = "Transformacional",
  "Carismático(a): exerce influência por meio de sua personalidade e visão inspiradora." = "Carismatico",
  "Nenhuma das anteriores" = "NDA"
)

# Aplicar a função recode para converter os valores em Estilo_de_lideranca
dados_tratados <- dados_tratados %>%
  mutate(Estilo_de_lideranca = recode(Estilo_de_lideranca, !!!correspondencia_lideranca))

# Definir vetor de correspondência para os novos rótulos de Qualidade_da_infraestrutura
correspondencia_infraestrutura <- c(
  "Péssima: a infraestrutura tecnológica é inoperante ou extremamente limitada, impossibilitando o trabalho dos colaboradores de forma eficiente e segura." = 1,
  "Ruim: a infraestrutura tecnológica apresenta problemas frequentes, com baixa confiabilidade e disponibilidade de recursos, e segurança precária." = 2,
  "Regular: a infraestrutura tecnológica apresenta algumas falhas, com confiabilidade e disponibilidade de recursos médias, e segurança que precisa de melhorias." = 3,
  "Boa: a infraestrutura tecnológica atende às necessidades básicas dos colaboradores, com confiabilidade e disponibilidade de recursos satisfatórias, e segurança adequada." = 4,
  "Excelente: a infraestrutura tecnológica atende plenamente às necessidades dos colaboradores, com alta confiabilidade, disponibilidade de recursos e segurança robusta." = 5
)

# Aplicar a função recode para converter os valores em Qualidade_da_infraestrutura
dados_tratados <- dados_tratados %>%
  mutate(Qualidade_da_infraestrutura = recode(Qualidade_da_infraestrutura, !!!correspondencia_infraestrutura))

# Transformar a variável "Faixa_etaria" em variável numérica de faixa
dados_tratados <- dados_tratados %>%
  mutate(Faixa_etaria = case_when(
    Faixa_etaria == "17-20 anos" ~ 1,
    Faixa_etaria == "21-25 anos" ~ 2,
    Faixa_etaria == "26-30 anos" ~ 3,
    Faixa_etaria == "31-40 anos" ~ 4,
    Faixa_etaria == "41-55 anos" ~ 5,
    Faixa_etaria == "56 anos ou mais" ~ 6,
    TRUE ~ NA_real_
  ))

# Transformar a variável "Faixa_salarial" em variável numérica de faixa
dados_tratados <- dados_tratados %>%
  mutate(Faixa_salarial = case_when(
    Faixa_salarial == "R$1.000-R$2.500;" ~ 1,
    Faixa_salarial == "R$2.501-R$4.000;" ~ 1,
    Faixa_salarial == "R$4.001-R$5.500;" ~ 2,
    Faixa_salarial == "R$5.501-R$8.000;" ~ 2,
    Faixa_salarial == "R$8.001-R$12.000;" ~ 3,
    Faixa_salarial == "Mais de R$12.000" ~ 3,
    TRUE ~ NA_real_
  ))

# Transformar a variável "Ferramentas_para_automatizacoes" em variável numérica de faixa
dados_tratados <- dados_tratados %>%
  mutate(Ferramentas_para_automatizacoes = case_when(
    Ferramentas_para_automatizacoes == "Básico" ~ 1,
    Ferramentas_para_automatizacoes == "Intemediário" ~ 2,
    Ferramentas_para_automatizacoes == "Avançado" ~ 3,
    TRUE ~ NA_real_
  ))

# Renomear as áreas de atuação
dados_tratados <- dados_tratados %>%
  mutate(Area_atuacao = case_when(
    # Grupo 1: Administrativo e Finanças
    Area_atuacao %in% c("Administrativo", "Planejamento Financeiro", "RH;", "Finanças", "Financeiro", 
                        "Jurídico", "Assistente") ~ "Finan.",
    
    # Grupo 2: Engenharia e Produção
    Area_atuacao %in% c("Engenharia", "Processos e Produção", "Operação. Usinagem", "Manutenção", 
                        "Operação Manutenção", "Engenharia de Desenvolvimento", "Produção", 
                        "Administrativo operacional/produção", "Melhoria Contínua") ~ "Eng.",
    
    # Grupo 3: Tecnologia e Análise de Dados
    Area_atuacao %in% c("Tecnologia da Informação", "Produtos Digitais", "Produtos digitais", "Análise de Dados;", 
                        "Governança de Dados", "Programador", "Gestão de produto e Tecnologia", 
                        "Gestao de produtos", "Desenvolvimento de Projetos", "Desenvolvimento de projetos", "Gestão de Produto e Tecnologia","Tecnologia", 
                        "Analista de negócios","T.I", "product manager", "Cx") ~ "Tec.",
    
    # Grupo 4: Educação e Consultoria
    Area_atuacao %in% c("Ensino Superior", "Educação Corporativa", "Coordenação Pedagógica", 
                        "Gestão de Projetos", "Consultoria", "sou professora de inglês autonoma", 
                        "Laboratório;", "Análise de dados e laboratório") ~ "Educ.",
    
    # Grupo 5: Marketing, Vendas e Logística
    Area_atuacao %in% c("Vendas", "Marketing", "Comercial", "Compras;","Produto", "Supply Chain", "Supply", "Logística","Vendas;") ~ "Com.",
    
    # Grupo 6: Saúde, Segurança e Sustentabilidade
    Area_atuacao %in% c("EHSQ - Segurança do Trabalho, Meio Ambiente, Saúde e Qualidade;", 
                        "Sustentabilidade de produtos","Saúde", "Sustentabilidade", "Gestão de saude") ~ "Saúde",
    
    TRUE ~ Area_atuacao
  ))

# Ajustar a variável Setor_empresa no dataframe dados_tratados
dados_tratados <- dados_tratados %>%
  mutate(Setor_empresa = case_when(
    Setor_empresa %in% c("Prestação de serviços;", "Comércio;", "Varejo", "Construção civil", 
                         "ONG", "Comércio e Prestação de serviços") ~ "Serv_Com",
    Setor_empresa %in% c("Educação", "Pesquisa científica", "educação", "Pesquisa e Desenvolvimento", "Pesquisa", "Poder Público", "Federal - Serviço Ambiental", "Serviço publico", "Setor Público") ~ "Educ_e_Gov",
    Setor_empresa %in% c("Tecnologia", "Engenharia", "Inovação aberta", "SaaS", "Software de Gestão", 
                         "Redes sociais", "Tecnologia e Inovação") ~ "Tec_Inovação",
    Setor_empresa %in% c("Financeiro", "saúde", "Jurídico", "Financeira", "Finanças", "Consultoria", "Marketing") ~ "Fin_Jurí",
    Setor_empresa %in% c("Indústria;", "Energia", "Operacional") ~ "Ind_Operacional",
    TRUE ~ NA_character_
  ))
# _______________________________________  Transformação de Variável alvo em binário

# Transformar variável em binária: Nunca, Raramente, Às vezes = 0; Frequentemente, Sempre = 1
dados_tratados$Implementar_planilhas_automatizadas <- recode(dados_tratados$Implementar_planilhas_automatizadas,
                                                              "Nunca" = 0,
                                                              "Raramente" = 0,
                                                              "Às vezes" = 0,
                                                              "Frequentemente" = 1,
                                                              "Sempre" = 1
)

# Transformar variável em binária: Nunca, Raramente, Às vezes = 0; Frequentemente, Sempre = 1
dados_tratados$Investimento_tecnologia <- recode(dados_tratados$Investimento_tecnologia,
                                                             "Nunca" = 0,
                                                             "Raramente" = 0,
                                                             "Às vezes" = 0,
                                                             "Frequentemente" = 1,
                                                             "Sempre" = 1
)

# Transformar a variável em binária: Nunca, Raramente, Às vezes = 0; Frequentemente, Sempre = 1
dados_tratados$capaz_implementar_planilhas <- recode(dados_tratados$capaz_implementar_planilhas,
                                                             "Nunca" = 0,
                                                             "Raramente" = 0,
                                                             "Às vezes" = 0,
                                                             "Frequentemente" = 1,
                                                             "Sempre" = 1
)

# _______________________________________  Filtro de Funções Administrativas 

# Filtrar os dados onde Funcoes_administrativas é igual a "Sim"
dados_filtrados <- dados_tratados %>%
  filter(Funcoes_administrativas == "Sim")
# Filtrar os dados onde capaz_implementar_planilhas é igual a "1"
#dados_filtrados <- dados_tratados %>%
 # filter(capaz_implementar_planilhas == "0")


# _______________________________________ Análise Exploratória de Dados (EDA)

# Verificar a estrutura dos dados filtrados
str(dados_filtrados)

# Resumo estatístico das variáveis numéricas nos dados filtrados
summary(dados_filtrados)

# Verificar a presença de valores ausentes nos dados filtrados
missing_values_filtrados <- sapply(dados_filtrados, function(x) sum(is.na(x)))
missing_values_filtrados

# Distribuição das variáveis categóricas nos dados filtrados
categorical_vars <- c("Funcoes_administrativas", "Area_atuacao", "Nivel_hierarquico", "Setor_empresa",
                      "Empresa_Publica_Privada", "Sexo", "Faixa_etaria", "Faixa_salarial", "Estilo_de_lideranca")

for (var in categorical_vars) {
  cat("Distribuição de", var, ":\n")
  print(table(dados_filtrados[[var]]))
  cat("\n")
}

# Definir as variáveis categóricas
categorical_vars <- c("Funcoes_administrativas", "Area_atuacao", "Nivel_hierarquico", "Setor_empresa",
                      "Empresa_Publica_Privada", "Sexo", "Faixa_etaria", "Faixa_salarial", "Estilo_de_lideranca")

# Criar um workbook
wb <- createWorkbook()

# Adicionar uma aba para cada variável categórica e preencher com dados de distribuição
for (var in categorical_vars) {
  distrib <- table(dados_filtrados[[var]])
  addWorksheet(wb, var)
  writeData(wb, var, as.data.frame(distrib), colNames = TRUE)
}

# Salvar o workbook em um arquivo Excel
saveWorkbook(wb, "distribuicao_variaveis_categoricas.xlsx", overwrite = TRUE)

#________________________________________ Random Florest 

library(randomForest)

# Verificar a variável alvo
na_count <- sum(is.na(dados_filtrados$capaz_implementar_planilhas))
cat("Número de valores NA na variável alvo:", na_count, "\n")

# Remover linhas com valores NA na variável alvo e em todas as preditoras
dados_filtrados <- dados_filtrados[complete.cases(dados_filtrados), ]

# Certificar-se de que a variável dependente é um fator
dados_filtrados$capaz_implementar_planilhas <- as.factor(dados_filtrados$capaz_implementar_planilhas)

# Verificar se a variável alvo tem mais de uma classe
num_classes <- length(unique(dados_filtrados$capaz_implementar_planilhas))
cat("Número de classes na variável alvo:", num_classes, "\n")

if (num_classes < 2) {
  stop("A variável alvo precisa ter pelo menos duas classes distintas para criar um modelo de classificação.")
}

# Definir a fórmula com as variáveis preditoras selecionadas
formula <- capaz_implementar_planilhas ~ Ferramentas_para_automatizacoes + Investimento_tecnologia + Empresa_agil
# Verificar se existem valores faltantes e remover se necessário
dados_filtrados <- na.omit(dados_filtrados)

# Definir os parâmetros do modelo Random Forest
ntree_opt <- 500
mtry_opt <- sqrt(ncol(dados_filtrados) - 1) # Número de preditores a serem considerados em cada split

# Construir o modelo Random Forest
modelo_rf <- randomForest(formula, data = dados_filtrados, ntree = ntree_opt, mtry = mtry_opt, importance = TRUE)

# Exibir o modelo
print(modelo_rf)

# Mostrar a importância das variáveis
importance_values <- importance(modelo_rf)
print(importance_values)

# Plotar a importância das variáveis
varImpPlot(modelo_rf)

# Mostrar os parâmetros otimizados
cat("Parâmetros otimizados:\n")
cat("Número de árvores (ntree):", ntree_opt, "\n")
cat("Número de variáveis por split (mtry):", mtry_opt, "\n")

#_________________________________ Árvore de Decisão 

# Verificar a variável alvo
na_count <- sum(is.na(dados_filtrados$capaz_implementar_planilhas))
cat("Número de valores NA na variável alvo:", na_count, "\n")

# Remover linhas com valores NA na variável alvo
dados_filtrados <- dados_filtrados %>% filter(!is.na(capaz_implementar_planilhas))

# Certificar-se de que a variável dependente é um fator
dados_filtrados$capaz_implementar_planilhas <- as.factor(dados_filtrados$capaz_implementar_planilhas)

# Verificar se a variável alvo tem mais de uma classe
num_classes <- length(unique(dados_filtrados$capaz_implementar_planilhas))
cat("Número de classes na variável alvo:", num_classes, "\n")

if (num_classes < 2) {
  stop("A variável alvo precisa ter pelo menos duas classes distintas para criar um modelo de classificação.")
}

# Definir a fórmula com as variáveis preditoras selecionadas
formula <- capaz_implementar_planilhas ~  Ferramentas_para_automatizacoes + Investimento_tecnologia + Empresa_agil
# Definir os parâmetros de controle para a árvore de decisão
control <- rpart.control(minsplit = 20, minbucket = 7, cp = 0.01)

# Construir o modelo de árvore de decisão
modelo_arvore <- rpart(formula, data = dados_filtrados, method = "class", control = control)

# Verificar os valores únicos na variável alvo
unique(dados_filtrados$capaz_implementar_planilhas)

# Definir o tamanho da figura e ajustar as margens
par(mar = c(4, 4, 4, 4))  # margens: inferior, esquerda, superior, direita

# Plotar árvore de decisão com ajustes
rpart.plot(
  modelo_arvore, 
  extra = 104,             # Mostrar a classe majoritária e a porcentagem de observações
  branch = 0.5, 
  under = TRUE, 
  type = 3,
  fallen.leaves = TRUE, 
  shadow.col = "gray",
  digits = 2,              # Mostrar porcentagens com duas casas decimais
  box.palette = "Blues",   # Colorir os nós de acordo com a classe
  compress = TRUE,         # Ajustar a compressão da árvore
  uniform = TRUE,          # Tornar os níveis mais uniformes
  split.cex = 1.0,         # Tamanho do texto das divisões
  split.fun = function(x, labs, digits, varlen, faclen) {
    # Função personalizada para formatar as divisões
    gsub("<= ", "<=", gsub("> ", ">", labs))
  },
  cex = 0.5               # Tamanho do texto dos rótulos dos nós
)



