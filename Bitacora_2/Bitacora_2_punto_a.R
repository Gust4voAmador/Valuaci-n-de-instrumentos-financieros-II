# Bitácora 2 · Punto (a) · Gustavo Amador Fonseca
# Ejecutar desde el proyecto que contiene Data/. Paquetes: xml2 y readxl.

# 1. Datos de la captura ----
carpeta_proyecto <- "."  # Cambiar si Data esta en otra carpeta del proyecto.
carpeta_data <- file.path(carpeta_proyecto, "Data")
archivo_excel <- file.path(carpeta_data, "NTR.NB.xlsx")

fecha_valoracion <- as.Date("2026-11-02")
fecha_vencimiento <- as.Date("2026-12-18")
hora_actualizacion_opciones <- "13:42:38"
zona_horaria_pantalla <- NA_character_ # No se identifica en las capturas.
S0 <- 69.925  # Last de la accion en la primera captura; no es precio de cierre.
moneda <- "USD"
fuente <- "LSEG/Refinitiv Workspace, pantalla Options OPW"

registro_capturas <- data.frame(
  captura = c("Primera", "Segunda"),
  fecha = rep(fecha_valoracion, 2),
  hora_accion = rep("13:42:24", 2),
  hora_opciones = rep(hora_actualizacion_opciones, 2),
  zona_horaria = rep(zona_horaria_pantalla, 2),
  accion_bid = c(69.91, 69.91),
  accion_ask = c(69.94, 69.95),
  accion_last = c(69.925, 69.900),
  moneda = moneda,
  stringsAsFactors = FALSE
)

# 2. Cadena de opciones y selección de la call más cercana al spot ----
cadena_opciones <- data.frame(
  strike = c(57.5, 60, 62.5, 65, 67.5, 70, 72.5, 75, 77.5, 80, 82.5),
  call_bid = c(12.40, 10.20, 8.10, 6.50, 4.70, 3.30, 2.35, 1.45, 0.90, 0.60, 0.35),
  call_ask = c(13.60, 11.20, 9.00, 7.00, 5.30, 3.90, 2.75, 1.90, 1.35, 0.90, 0.65),
  call_last = c(12.45, 10.25, 8.10, 6.25, 4.65, 3.54, 2.50, 1.69, 1.00, 0.75, 0.53),
  call_volumen = c(NA, NA, NA, NA, NA, NA, 1, 7, 2, 50, NA),
  call_interes_abierto = c(5, 225, 890, 171, 171, 193, 99, 1608, 168, 1429, 160),
  call_cambio_interes_abierto = c(0, 0, 0, 0, 0, 0, 0, 10, 0, -1, 0),
  call_iv_pct = c(39.7166, 37.6474, 36.3889, 37.4871, 36.0185, 35.5147,
                  35.7880, 35.1369, 35.4989, 35.9564, 36.5838),
  call_delta = c(0.9355, 0.8975, 0.8375, 0.7456, 0.6490, 0.5371,
                 0.4255, 0.3190, 0.2335, 0.1672, 0.1182),
  put_bid = c(0.20, 0.40, 0.75, 1.25, 2.05, 3.20, 4.50, 6.20, 8.10, 10.20, 12.40),
  put_ask = c(0.40, 0.55, 0.95, 1.60, 2.45, 3.50, 5.00, 6.70, 8.60, 11.10, 13.30),
  put_last = c(0.35, 0.45, 0.90, 1.70, 2.45, 3.40, 5.15, 6.95, 8.85, 10.37, 13.35),
  put_volumen = c(22, NA, NA, NA, NA, 6, NA, NA, NA, 50, NA),
  put_interes_abierto = c(48, 115, 160, 211, 104, 283, 66, 81, 95, 8, 100),
  put_cambio_interes_abierto = c(0, 0, 0, 1, 1, -5, 0, 0, 0, 0, 0),
  put_iv_pct = c(39.9005, 37.4855, 36.9129, 36.3969, 35.9204, 35.4363,
                 35.1274, 35.1919, 34.7615, 37.5424, 37.1706),
  put_delta = c(-0.0658, -0.1023, -0.1662, -0.2510, -0.3539, -0.4678,
                -0.5836, -0.6904, -0.7859, -0.8347, -0.8966)
)

cadena_opciones$call_mid <- with(cadena_opciones, (call_bid + call_ask) / 2)
cadena_opciones$put_mid <- with(cadena_opciones, (put_bid + put_ask) / 2)
cadena_opciones$call_spread <- with(cadena_opciones, call_ask - call_bid)
cadena_opciones$put_spread <- with(cadena_opciones, put_ask - put_bid)

# Agrupa las cotizaciones de un tipo de opción y añade los datos del contrato.
armar_opciones <- function(tipo, prefijo) {
  columnas <- c("bid", "ask", "last", "volumen", "interes_abierto",
                "cambio_interes_abierto", "iv_pct", "delta", "mid", "spread")
  resultado <- cadena_opciones[paste0(prefijo, "_", columnas)]
  names(resultado) <- columnas
  resultado <- cbind(
    data.frame(
      subyacente = "NTR",
      ric_accion_pantalla = "NTR.NB",
      contrato = NA_character_, # El identificador de cada contrato no es visible.
      tipo = tipo,
      strike = cadena_opciones$strike,
      fecha_captura = fecha_valoracion,
      hora_actualizacion_opciones = hora_actualizacion_opciones,
      zona_horaria_pantalla = zona_horaria_pantalla,
      vencimiento = fecha_vencimiento,
      moneda = moneda,
      S0 = S0,
      fuente = fuente,
      stringsAsFactors = FALSE
    ), resultado
  )
  resultado$iv <- resultado$iv_pct / 100
  resultado$spread_relativo <- resultado$spread / resultado$mid
  resultado$dias_al_vencimiento <- as.integer(fecha_vencimiento - fecha_valoracion)
  resultado$T <- resultado$dias_al_vencimiento / 365
  # Last puede ser una transaccion antigua: se registra sin reemplazarlo por mid.
  resultado$last_fuera_bid_ask <- with(resultado, last < bid | last > ask)
  resultado
}

opciones <- rbind(armar_opciones("Call", "call"), armar_opciones("Put", "put"))
rownames(opciones) <- NULL
calls <- opciones[opciones$tipo == "Call", ]
puts <- opciones[opciones$tipo == "Put", ]

# Todos los pares se conservan; la seleccion de calidad se hara al analizarlos.
pares_opciones <- cadena_opciones[
  c("strike", "call_bid", "call_ask", "call_mid", "put_bid", "put_ask", "put_mid")
]
pares_opciones$fecha_captura <- fecha_valoracion
pares_opciones$vencimiento <- fecha_vencimiento
pares_opciones$S0 <- S0
pares_opciones$T <- as.integer(fecha_vencimiento - fecha_valoracion) / 365

call_principal <- calls[which.min(abs(calls$strike - S0)), ]

stopifnot(
  nrow(cadena_opciones) == 11L,
  !anyDuplicated(cadena_opciones$strike),
  nrow(opciones) == 22L,
  all(opciones$ask >= opciones$bid),
  all(opciones$bid > 0),
  all(opciones$T > 0)
)
print(call_principal[c("tipo", "strike", "bid", "ask", "mid", "T")], row.names = FALSE)

# 3. Leer y ordenar los cierres anteriores a la captura ----
# Esta exportación guarda la fecha como ISO 8601 en el XML del Excel.
# Leemos ambas celdas de cada fila para conservar el par fecha-cierre.
if (!requireNamespace("xml2", quietly = TRUE)) {
  stop("Instalá xml2 con install.packages('xml2') y volvé a ejecutar el script.")
}
miembro <- "xl/worksheets/sheet1.xml"  # Hoja Table Data de NTR.NB.xlsx
# Lee el XML de la hoja que contiene el histórico de NTR.
leer_hoja <- function(archivo, miembro) {
  conexion <- unz(archivo, miembro, open = "rb")
  on.exit(close(conexion))
  tamano <- unzip(archivo, list = TRUE)
  bytes <- tamano$Length[match(miembro, tamano$Name)]
  xml2::read_xml(readBin(conexion, "raw", n = bytes))
}
hoja <- leer_hoja(archivo_excel, miembro)
xml2::xml_ns_strip(hoja)
filas <- xml2::xml_find_all(hoja, ".//sheetData/row[number(@r) >= 3]")

# Extrae una celda por fila, conservando el orden de los cierres.
valor_columna <- function(columna) {
  ruta <- sprintf("./c[starts-with(@r, '%s')]/v", columna)
  xml2::xml_text(xml2::xml_find_first(filas, ruta))
}
fecha_texto <- valor_columna("A")
historico <- data.frame(
  fecha = as.Date(substr(fecha_texto, 1, 10), format = "%Y-%m-%d"),
  cierre = as.numeric(valor_columna("B"))
)
if (!nrow(historico) || anyNA(historico$fecha)) {
  stop("Revisá la hoja Table Data de NTR.NB.xlsx. Se esperaban fechas ISO en A desde la fila 3.")
}
historico <- historico[historico$fecha < fecha_valoracion, ]
historico <- historico[order(historico$fecha), ]
if (nrow(historico) < 3L) {
  stop(sprintf("Solo hay %d cierres anteriores al %s en %s.",
               nrow(historico), fecha_valoracion, archivo_excel))
}
stopifnot(!anyDuplicated(historico$fecha),
          all(is.finite(historico$cierre)), all(historico$cierre > 0))


# 4. Estimar volatilidad, deriva y su intervalo histórico ----
# Rₜ = ln(Sₜ / Sₜ₋₁).
retornos <- diff(log(historico$cierre))
n <- length(retornos)
sesiones <- 252
# σ̂ = s(R) √252, donde s(R) es la desviación estándar diaria.
sigma <- sd(retornos) * sqrt(sesiones)
# μ̂ = 252 R̄ + σ̂²/2.
mu <- mean(retornos) * sesiones + sigma^2 / 2
# Error estándar aproximado bajo retornos normales independientes.
# La media y la varianza muestrales son independientes bajo normalidad.
# EE(μ̂) ≈ √[252 σ̂²/n + σ̂⁴/(2(n−1))].
ee_mu <- sqrt(sesiones * sigma^2 / n + sigma^4 / (2 * (n - 1)))
# IC del 95 % para μ: μ̂ ± 1.96 EE(μ̂).
ic_mu <- mu + c(-1, 1) * qnorm(0.975) * ee_mu
# Formatea las tasas anuales como porcentajes para la consola.
pct <- function(x) sprintf("%.2f%%", 100 * x)
cat(sprintf("%d retornos | Volatilidad: %s | Deriva: %s | IC 95 %%: [%s, %s]\n",
            n, pct(sigma), pct(mu), pct(ic_mu[1]), pct(ic_mu[2])))

# 5. Tasa libre de riesgo y plazo de la call ----
# DTB4WK (30/10/2026): convertir descuento bancario a tasa continua.
d <- 3.91 / 100
# r = −(365/28) ln(1 − d × 28/360).
r <- -log1p(-d * 28 / 360) * 365 / 28
q <- 0  # Sin dividendos durante la vida de la call.
K <- call_principal$strike
# T = días al vencimiento / 365.
T <- call_principal$T

# Contraste con la letra específica del Excel, que vence el 17 de diciembre.
archivo_letra <- file.path(carpeta_data, "Price History_US912797VH74(1).xlsx")
letra <- readxl::read_excel(archivo_letra, range = "A18:D122",
                            col_types = c("numeric", "skip", "numeric", "numeric"))
# Las fechas de esta letra son números de serie de Excel.
letra$Date <- as.Date(letra$Date, origin = "1899-12-30")
fila <- letra[which(as.Date(letra$Date) == as.Date("2026-10-30")), ]
stopifnot(nrow(fila) == 1, is.finite(fila$Bid), fila$Ask >= fila$Bid)
precio_letra <- (fila$Bid + fila$Ask) / 2
# Aproximación ACT/365 desde la fecha de cotización (sin ajuste de liquidación).
dias_letra <- as.numeric(as.Date("2026-12-17") - as.Date(fila$Date))
# r = −ln(precio medio / 100) × 365 / días al vencimiento.
r_letra <- -log(precio_letra / 100) * 365 / dias_letra
cat("Tasa usada:", pct(r), "| Contraste con la letra de Refinitiv:", pct(r_letra), "\n")

# 6. Simular la call bajo P y repesar los pagos ----
set.seed(20261001)  # Semilla del laboratorio.
N <- 500000L
# W_T = √T Z, con Z ~ N(0, 1).
W_T <- rnorm(N, sd = sqrt(T))

# Devuelve el precio medio y su intervalo de Monte Carlo al 95 %.
ic_mc <- function(aportes) {
  media <- mean(aportes)
  # IC de Monte Carlo: Ȳ ± 1.96 s(Y)/√N, donde Y es el aporte por simulación.
  margen <- qnorm(0.975) * sd(aportes) / sqrt(length(aportes))
  c(precio = media, inferior = media - margen, superior = media + margen)
}

# Para una deriva, calcula ambos precios con los mismos movimientos simulados.
valorar <- function(deriva) {
  # λ = (μ + q − r)/σ. Aquí μ es la deriva elegida para este escenario.
  lambda <- (deriva + q - r) / sigma
  # S_T = S₀ exp[(μ − σ²/2)T + σW_T], bajo ℙ.
  S_T <- S0 * exp((deriva - sigma^2 / 2) * T + sigma * W_T)
  # L_T cambia el peso de cada pago al pasar de P a Q.
  # L_T = exp(−λW_T − λ²T/2).
  L_T <- exp(-lambda * W_T - lambda^2 * T / 2)
  # X = máx(S_T − K, 0). Su valor descontado en este sorteo es exp(−rT) X.
  pago_desc <- exp(-r * T) * pmax(S_T - K, 0)
  # Repesado: exp(−rT) Eℙ[L_T X]. Ingenuo: exp(−rT) Eℙ[X].
  tabla <- rbind(Repesado = ic_mc(L_T * pago_desc),
                 Ingenuo = ic_mc(pago_desc))
  data.frame(mu = deriva, lambda = lambda, media_peso = mean(L_T),
             metodo = rownames(tabla), tabla, row.names = NULL)
}

derivas <- c(ic_mu[1], mu, ic_mu[2])
resultados <- do.call(rbind, lapply(derivas, valorar))
escenario <- rep(c("Extremo inferior", "Estimación", "Extremo superior"), each = 2)
resultados <- data.frame(escenario, resultados)
print(resultados, digits = 5, row.names = FALSE)

# 7. Contrastar con Black–Scholes e interpretar la prima de riesgo ----
# d₁ = [ln(S₀/K) + (r − q + σ²/2)T] / (σ√T).
d1 <- (log(S0 / K) + (r - q + sigma^2 / 2) * T) / (sigma * sqrt(T))
# d₂ = d₁ − σ√T.
d2 <- d1 - sigma * sqrt(T)
# C₀ = S₀ exp(−qT) Φ(d₁) − K exp(−rT) Φ(d₂), con Φ la normal estándar acumulada.
referencia <- S0 * exp(-q * T) * pnorm(d1) - K * exp(-r * T) * pnorm(d2)
# Prima = μ + q − r. Sumamos q − r a ambos extremos del IC de μ.
ic_prima <- ic_mu + q - r
cat(sprintf("Black–Scholes: USD %.3f | Punto medio observado: USD %.3f\n",
            referencia, call_principal$mid))
cat(sprintf("Prima de riesgo: %s | IC 95 %%: [%s, %s]\n",
            pct(mu + q - r), pct(ic_prima[1]), pct(ic_prima[2])))
if (ic_prima[1] <= 0 && ic_prima[2] >= 0) {
  cat("El intervalo incluye cero: la muestra no permite establecer el signo de la prima.\n")
}
