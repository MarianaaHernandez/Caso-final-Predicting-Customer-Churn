# Caso: Predicción y Prevención de Deserción de Clientes
# Empresa: QWE INC.

library(readxl)
library(tidyverse)
library(ggplot2)
library(flextable)
library(officer)
library(broom)
library(caret)
library(scales)
library(dplyr)
library(gt)

# ============================================================
# Cargar datos
# ============================================================
library(readxl)
datos <- read_excel('/Users/maru/Desktop/Javeriana/4/Analitica de los negocios/business-analytics-2026-1/Cases/Predicting _Customer_Churn/DATA.xlsx',
                    sheet = "Case Data",
                    col_names = c(
                      "ID", "Customer_Age", "Churn", "CHI_Month0", "CHI_01",
                      "Support_Month0", "Support_01", "SP_Month0", "SP_01",
                      "Logins_01", "Blog_01", "Views_01", "Days_Login_01"
                    ),
                    skip = 1)

# Renombrar columnas para facilitar el trabajo en R
names(datos) <- c(
  "id", "edad_cliente", "churn",
  "chi_mes0", "chi_cambio",
  "casos_soporte_mes0", "casos_soporte_cambio",
  "prioridad_mes0", "prioridad_cambio",
  "logins_cambio", "articulos_cambio",
  "views_cambio", "dias_sin_login_cambio"
)

cat("Dimensiones:", nrow(datos), "filas x", ncol(datos), "columnas\n")
cat("Clientes que desertaron:", sum(datos$churn), "(", round(mean(datos$churn)*100, 1), "%)\n")

# ============================================================
# Tabla descriptiva
# ============================================================
vars_desc <- c(
  "edad_cliente", "chi_mes0", "chi_cambio",
  "casos_soporte_mes0", "casos_soporte_cambio",
  "logins_cambio", "views_cambio", "dias_sin_login_cambio"
)

nombres_legibles <- c(
  "Edad del cliente (meses)", "CHI Score (mes 0)", "Cambio CHI Score",
  "Casos de soporte (mes 0)", "Cambio casos de soporte",
  "Cambio en logins", "Cambio en vistas", "Cambio días sin login"
)

tabla_desc <- data.frame(
  Variable  = nombres_legibles,
  N         = sapply(datos[vars_desc], function(x) sum(!is.na(x))),
  Media     = sapply(datos[vars_desc], function(x) round(mean(x, na.rm = TRUE), 2)),
  Mediana   = sapply(datos[vars_desc], function(x) round(median(x, na.rm = TRUE), 2)),
  SD        = sapply(datos[vars_desc], function(x) round(sd(x, na.rm = TRUE), 2)),
  Minimo    = sapply(datos[vars_desc], function(x) round(min(x, na.rm = TRUE), 2)),
  Maximo    = sapply(datos[vars_desc], function(x) round(max(x, na.rm = TRUE), 2)),
  row.names = NULL
)

print(tabla_desc)

tabla_desc %>%
  gt() %>%
  tab_header(
    title = "Tabla Descriptiva de Variables",
  ) %>%
  cols_label(
    Variable = "Variable",
    N = "N",
    Media = "Media",
    Mediana = "Mediana",
    SD = "Desv. Estándar",
    Minimo = "Mín",
    Maximo = "Máx"
  ) %>%
  fmt_number(
    columns = c(Media, Mediana, SD, Minimo, Maximo),
    decimals = 2
  ) %>%
  cols_align(
    align = "center",
    -Variable
  )


# ============================================================
# Modelo de regresión Logit (Variable binaria)
# ============================================================
# Como grupo decidimos usar logit ya que la variable dependiente es binaria (0/1)
modelo_logit <- glm(churn ~ edad_cliente + chi_mes0 + chi_cambio +
                     casos_soporte_mes0 + casos_soporte_cambio +
                     logins_cambio + views_cambio + dias_sin_login_cambio,
                   data = datos, family = binomial(link = "logit"))
summary(modelo_logit)
# R^2 
r2_mcfadden <- 1 - (logLik(modelo_logit) / logLik(update(modelo_logit, . ~ 1)))
round(r2_mcfadden, 4)

# Creacion tabla de coeficientes para exportar a word
tabla_coef <- tidy(modelo_logit) %>%
  mutate(
    Variable = c("Intercepto", "Edad", "CHI 0", "Δ CHI",
                 "Soporte 0", "Δ Soporte", "Δ Logins",
                 "Δ Vistas", "Δ Días sin login"),
    Coef = round(estimate, 4),
    p = round(p.value, 4)
  ) %>%
  select(Variable, Coef, p)

ft_modelo <- flextable(tabla_coef) %>%
  bold(part = "header") %>%
  autofit()

# Documento Word con los resultados
doc <- read_docx() %>%
  body_add_par("Resultados modelo logit", style = "heading 1") %>%
  body_add_flextable(ft_modelo) %>%
  body_add_par(
    paste0("McFadden R² = ", round(r2_mcfadden, 4)),
    style = "Normal"
  )

print(doc, target = "Resultados_QWE.docx")

# ============================================================
# Matriz de Confusión
# ============================================================
datos$prob_predicha <- predict(modelo, type = "response")

# Calcular umbral óptimo (Youden: maximiza sensibilidad + especificidad)
# Necesario porque solo el 5% deserta, con 0.5 el modelo predice siempre "no deserta"
umbrales <- seq(0.03, 0.4, by = 0.005)
resultados_umbral <- data.frame(umbral = umbrales, youden = NA)
for (i in seq_along(umbrales)) {
  u <- umbrales[i]
  pred_tmp <- ifelse(datos$prob_predicha >= u, 1, 0)
  tp <- sum(pred_tmp == 1 & datos$churn == 1)
  fp <- sum(pred_tmp == 1 & datos$churn == 0)
  tn <- sum(pred_tmp == 0 & datos$churn == 0)
  fn <- sum(pred_tmp == 0 & datos$churn == 1)
  sens <- if ((tp + fn) > 0) tp / (tp + fn) else 0
  spec <- if ((tn + fp) > 0) tn / (tn + fp) else 0
  resultados_umbral$youden[i] <- sens + spec - 1
}
umbral_optimo <- resultados_umbral$umbral[which.max(resultados_umbral$youden)]
cat("\nUmbral óptimo (Youden):", umbral_optimo, "\n")

datos$churn_predicho <- ifelse(datos$prob_predicha >= umbral_optimo, 1, 0)

# Calcular métricas
conf_mat <- confusionMatrix(
  factor(datos$churn_predicho, levels = c(0, 1)),
  factor(datos$churn,          levels = c(0, 1)),
  positive = "1"
)
cat("\nMétricas del modelo:\n")
print(conf_mat)

conf_df <- as.data.frame(conf_mat$table)
conf_df$Etiqueta <- paste0(conf_df$Freq)
conf_df$Pct <- paste0(round(conf_df$Freq / sum(conf_df$Freq) * 100, 1), "%")

conf_df <- conf_df %>%
  mutate(
    Prediccion_label = ifelse(Prediction == "1", "Predicho: Deserta", "Predicho: No Deserta"),
    Real_label       = ifelse(Reference   == "1", "Real: Deserta",     "Real: No Deserta")
  )

grafico_conf <- ggplot(conf_df, aes(x = Prediccion_label, y = Real_label, fill = Freq)) +
  geom_tile(color = "white", linewidth = 1.5) +
  geom_text(aes(label = paste0(Freq, "\n(", Pct, ")")),
            size = 7, fontface = "bold", color = "white") +
  scale_fill_gradient(low = "#3498db", high = "#1a5276",
                      name = "Frecuencia") +
  labs(
    title    = "Matriz de Confusión - Modelo Logit",
    subtitle = paste0("Umbral de clasificación: 0.5  |  Exactitud: ",
                      round(conf_mat$overall["Accuracy"] * 100, 1), "%"),
    x = "Valor Predicho",
    y = "Valor Real"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title    = element_text(face = "bold", size = 16, hjust = 0.5),
    plot.subtitle = element_text(size = 12, hjust = 0.5, color = "gray40"),
    axis.title    = element_text(face = "bold", size = 13),
    axis.text     = element_text(size = 12),
    legend.position = "right",
    panel.grid    = element_blank()
  )

grafico_conf
ggsave("matriz_confusion_QWE.png", grafico_conf, width = 8, height = 6, dpi = 150)
cat("Gráfico guardado: matriz_confusion_QWE.png\n")


