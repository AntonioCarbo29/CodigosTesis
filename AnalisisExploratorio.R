# ======================================================
# ANÁLISIS EXPLORATORIO
# - Imprime TODAS las opciones de TODAS las preguntas
# - Frecuencia y porcentaje sobre N total
# - Guarda Excel completo
# ======================================================

# NOTA (ya instaladas, no ejecutar):
# install.packages("readxl")
# install.packages("janitor")
# install.packages("dplyr")
# install.packages("stringr")
# install.packages("openxlsx")

library(readxl)
library(janitor)
library(dplyr)
library(stringr)
library(openxlsx)

# 1) Ruta del archivo
ruta_carpeta <- "C:/Users/Admin/Desktop/CARPETA TESIS FINAL  Antonio Carbo"
archivo_datos <- "DatosLimpiosAC.xlsx"
ruta_completa <- file.path(ruta_carpeta, archivo_datos)

# 2) Cargar datos
datos <- read_excel(ruta_completa) %>% clean_names()

# Total de encuestados
n_total <- nrow(datos)

# 3) Función: tabla COMPLETA por pregunta (sin resumir)
tabla_completa <- function(x, n_total) {

  x <- as.character(x)
  x <- str_squish(x)

  # NA o vacíos se consideran "Sin especificar"
  x[is.na(x) | x == ""] <- "Sin especificar"

  tab <- sort(table(x), decreasing = TRUE)

  data.frame(
    Respuesta  = names(tab),
    Frecuencia = as.integer(tab),
    Porcentaje = round(100 * as.integer(tab) / n_total, 1),
    check.names = FALSE
  )
}

# 4) IMPRIMIR TODO EN R (SIN OCULTAR NADA)
cat("\n====================================\n")
cat("ANÁLISIS EXPLORATORIO COMPLETO\n")
cat("Total encuestados (N):", n_total, "\n")
cat("====================================\n")

for (pregunta in names(datos)) {

  cat("\n====================================\n")
  cat("Pregunta:", pregunta, "\n")
  cat("====================================\n")

  resultado <- tabla_completa(datos[[pregunta]], n_total)
  print(resultado, row.names = FALSE)
}

# 5) EXPORTAR TODO A EXCEL (UNA HOJA POR PREGUNTA)
wb <- createWorkbook()

for (i in seq_along(names(datos))) {

  pregunta <- names(datos)[i]
  tabla <- tabla_completa(datos[[pregunta]], n_total)

  nombre_hoja <- paste0("Pregunta_", sprintf("%02d", i))
  addWorksheet(wb, nombre_hoja)

  # Escribir nombre completo de la pregunta
  writeData(
    wb,
    nombre_hoja,
    paste("Pregunta (variable):", pregunta),
    startRow = 1,
    startCol = 1
  )

  # Escribir tabla completa
  writeData(
    wb,
    nombre_hoja,
    tabla,
    startRow = 3,
    startCol = 1
  )

  setColWidths(wb, nombre_hoja, cols = 1:3, widths = "auto")
}

# 6) Guardar archivo final
salida <- file.path(ruta_carpeta, "Analisis_Exploratorio_Completo.xlsx")
saveWorkbook(wb, salida, overwrite = TRUE)

cat("\nArchivo generado correctamente:\n", salida, "\n")
