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

# ============================================================
# Cargar datos
# ============================================================
library(readxl)
datos <- read_excel("~/Documents/Universidad Javeriana/Semestre 4/Analítica de Datos/Casos/Caso 4/DATA.xlsx",
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
