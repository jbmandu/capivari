library(forecast)
library(dplyr)

# Supondo que df seja seu dataframe e a segunda coluna contém os dados de demanda
ts_demanda <- ts(df[[2]], start = c(2011, 1), end = c(2019, 12), frequency = 12)

# Inicializando o dataframe de resultados
resultados <- data.frame(Ano_Treino_Fim = integer(), MAPE = numeric(), Modelo = character())

# Definições iniciais
anos_treino_inicial <- 7
tamanho_passo_teste <- 2


total_anos <- 2019 - 2011 + 1
total_iteracoes <- total_anos - anos_treino_inicial

for (i in 0:(total_iteracoes - tamanho_passo_teste)) {
  ini_treino <- c(2011, 1)
  fim_treino <- c(2011 + i + anos_treino_inicial - 1, 12)
  ini_teste <- c(2011 + i + anos_treino_inicial, 1)
  fim_teste <- c(2011 + i + anos_treino_inicial + tamanho_passo_teste -1, 12)
  
  if(fim_teste[1] <= 2019) {
    dtreino <- window(ts_demanda, start = ini_treino, end = fim_treino)
    dteste <- window(ts_demanda, start = ini_teste, end = fim_teste)
    
    
    # Modelo SES
    modelo_ses <- ses(dtreino, h = tamanho_passo_teste * 12)
    forecast_ses <- forecast(modelo_ses, h = tamanho_passo_teste * 12)
    mape_ses <- accuracy(forecast_ses, dteste)[2, "MAPE"]
    resultados <- rbind(resultados, 
                        data.frame(
                          Ano_Treino_Fim = 2011 + i + anos_treino_inicial - 1, 
                          MAPE = mape_ses, 
                          Modelo = "SES"))
    
    
    # Modelo Holt-Winters Aditivo
    mod_hw_ad = hw(dtreino,h=tamanho_passo_teste * 12,seasonal="additive")
    forecast_hw_ad <- forecast(mod_hw_ad, h = tamanho_passo_teste * 12)
    mape_hw_ad <- accuracy(forecast_hw_ad, dteste)[2, "MAPE"]
    resultados <- rbind(resultados, 
                        data.frame(
                          Ano_Treino_Fim = 2011 + i + anos_treino_inicial - 1, 
                          MAPE = mape_hw_ad, 
                          Modelo = "HW-Ad"))
    
    
    # Modelo Holt-Winters Multiplicativo
    mod_hw_mu = hw(dtreino,h=tamanho_passo_teste * 12,seasonal="multiplicative")
    forecast_hw_mu <- forecast(mod_hw_mu, h = tamanho_passo_teste * 12)
    mape_hw_mu <- accuracy(forecast_hw_mu, dteste)[2, "MAPE"]
    resultados <- rbind(resultados, 
                        data.frame(
                          Ano_Treino_Fim = 2011 + i + anos_treino_inicial - 1, 
                          MAPE = mape_hw_mu, 
                          Modelo = "HW-Mu"))
    
    
    # Modelo ARIMA
    modelo_arima <- auto.arima(dtreino, trace = FALSE) 
    # trace=T pode ser alterado para trace=FALSE para evitar a saída excessiva
    
    forecast_arima <- forecast(modelo_arima, h = tamanho_passo_teste * 12)
    mape_arima <- accuracy(forecast_arima, dteste)[2, "MAPE"]
    resultados <- rbind(resultados, 
                        data.frame(
                          Ano_Treino_Fim = 2011 + i + anos_treino_inicial - 1, 
                          MAPE = mape_arima, 
                          Modelo = "ARIMA"))
    
    # Aqui você pode incluir testes de diagnóstico adicionais se necessário
    # Por exemplo, você pode incluir o teste de Ljung-Box nos resíduos do modelo ARIMA
    # lb_test <- Box.test(arima_model$residuals, lag = 20, type = "Ljung-Box")
    # print(paste("P-value do teste de Ljung-Box para ARIMA: ", lb_test$p.value))
    
    
    # Modelo ETS
    modelo_ets <- ets(dtreino, model = "ZZZ")
    forecast_ets <- forecast(modelo_ets, h = tamanho_passo_teste * 12)
    mape_ets <- accuracy(forecast_ets, dteste)[2, "MAPE"]
    resultados <- rbind(resultados, 
                        data.frame(
                          Ano_Treino_Fim = 2011 + i + anos_treino_inicial - 1, 
                          MAPE = mape_ets, 
                          Modelo = "ETS"))
    
    # Modelo ETS com transformação de Box-Cox
    l = BoxCox.lambda(ts_demanda)
    BOX_ETS=forecast(ets(dtreino, lambda = l),h=tamanho_passo_teste * 12)
    mape_ets_bc <- accuracy(BOX_ETS, dteste)[2, "MAPE"]
    resultados <- rbind(resultados, 
                        data.frame(
                          Ano_Treino_Fim = 2011 + i + anos_treino_inicial - 1, 
                          MAPE = mape_ets_bc, 
                          Modelo = "ETS-BC"))
    
    # Modelo NNAR - Neural Network AutoRegressive
    
    modelo_nnar <- nnetar(dtreino)
    forecast_nnar <- forecast(modelo_nnar, h = tamanho_passo_teste * 12)
    mape_nnar <- accuracy(forecast_nnar, dteste)[2, "MAPE"]
    resultados <- rbind(resultados, 
                        data.frame(
                          Ano_Treino_Fim = 2011 + i + anos_treino_inicial - 1, 
                          MAPE = mape_nnar, 
                          Modelo = "NNAR"))
 

    # Se quiser incluir outros testes de diagnóstico ou visualizações específicas, faça isso aqui
    
  } else {
    break # Interromper o loop se a janela de teste ultrapassar os dados disponíveis
  }
}

print(resultados)

## SALVANDO TABELA EM EXCEL ###

# Obter o caminho do diretório de trabalho atual
caminho <- getwd()

# Definir o caminho completo do arquivo Excel concatenando o ano inicial e o final

arquivo_xls <- paste0(caminho, "/Resultados-cross-val2-1.xlsx")

# Exportar dados para um arquivo Excel .xlsx
write_xlsx(resultados, arquivo_xls)

