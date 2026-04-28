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
libreary(dplyr)

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

# Tabla bonita 
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

