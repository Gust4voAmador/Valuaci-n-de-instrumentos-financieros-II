# Bitácora 2 · Partes (a) a (e)
# Gustavo Amador Fonseca C20451, Henri Gerard Gabert Hidalgo B93096
# Guardá este .R y los dos Excel en la misma carpeta y ejecutá desde el inicio.
# Paquetes: xml2 y readxl (rstudioapi permite ubicar el script en RStudio).

# 1. Datos de la captura ----
# Ubica el script al usar Source, Rscript o ejecutar líneas desde RStudio.
rutas_script <- unlist(lapply(sys.frames(), function(x) x$ofile))
argumento_script <- grep("^--file=", commandArgs(), value = TRUE)
if (length(argumento_script)) {
  rutas_script <- c(rutas_script, sub("^--file=", "", argumento_script))
}
if (requireNamespace("rstudioapi", quietly = TRUE) && rstudioapi::isAvailable()) {
  ruta_editor <- tryCatch(rstudioapi::getSourceEditorContext()$path,
                          error = function(e) "")
  rutas_script <- c(rutas_script, ruta_editor)
}
rutas_script <- rutas_script[!is.na(rutas_script) & nzchar(rutas_script)]
carpetas <- unique(c(dirname(rutas_script), getwd()))
carpetas <- carpetas[file.exists(file.path(carpetas, "NTR.NB.xlsx"))]
if (!length(carpetas)) {
  stop("No se encontró NTR.NB.xlsx junto al script. Guardá el .R y los Excel en la misma carpeta y ejecutá con Source.")
}
carpeta_proyecto <- normalizePath(carpetas[1], winslash = "/", mustWork = TRUE)
carpeta_data <- carpeta_proyecto
archivo_excel <- file.path(carpeta_data, "NTR.NB.xlsx")
archivo_letra <- file.path(carpeta_data, "Price History_US912797VH74.xlsx")
if (!file.exists(archivo_letra)) {
  stop("Falta Price History_US912797VH74.xlsx en la misma carpeta del script.")
}

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
fechas <- as.Date(substr(fecha_texto, 1, 10), format = "%Y-%m-%d")
# Si Excel guardó fechas como números de serie, convertirlas también.
seriales <- suppressWarnings(as.numeric(fecha_texto))
usar_serial <- is.na(fechas) & is.finite(seriales)
fechas[usar_serial] <- as.Date(seriales[usar_serial], origin = "1899-12-30")
historico <- data.frame(
  fecha = fechas,
  cierre = as.numeric(valor_columna("B"))
)
if (!nrow(historico) || anyNA(historico$fecha)) {
  stop("Revisá la hoja Table Data de NTR.NB.xlsx. Se esperaban fechas en A desde la fila 3.")
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

# PARTE (a). Cambio de medida y valoración por repesado ----

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


# PARTE (b). Un evento, tres medidas ----
# Reutilizamos S0, K, T, mu, sigma, r y q de (a).
# A = {S_T > K}: la call termina dentro del dinero.
# (S_T - K)+ = S_T 1_A - K 1_A.
# Con V_t = exp(qt) S_t como numerario (aquí q = 0, por lo que V = S):
# C0 = S0 exp(-qT) QV(A) - K exp(-rT) Q(A).

# 8. Simular el evento bajo P, Q y QV ----
set.seed(20261022)
n_b <- 500000L
z_b <- rnorm(n_b)  # Mismos sorteos para comparar las tres distribuciones.

# Cada simulación devuelve TRUE si el precio terminal supera el strike.
evento_b <- function(deriva) {
  precio_final <- S0 * exp((deriva - sigma^2 / 2) * T + sigma * sqrt(T) * z_b)
  precio_final > K
}

# P usa la deriva histórica; Q, la bancaria; QV añade sigma^2 al cambiar numerario.
derivas_b <- c(P = mu, Q = r - q, QV = r - q + sigma^2)
indicadores_b <- sapply(derivas_b, evento_b)
prob_mc_b <- colMeans(indicadores_b)  # Proporción de veces que ocurre A.

# 9. Comparar las probabilidades simuladas y teóricas ----
d2_b <- (log(S0 / K) + (r - q - sigma^2 / 2) * T) / (sigma * sqrt(T))
d1_b <- d2_b + sigma * sqrt(T)
dP_b <- (log(S0 / K) + (mu - sigma^2 / 2) * T) / (sigma * sqrt(T))
# Q(A) = Phi(d2), QV(A) = Phi(d1); bajo P se usa mu en lugar de r - q.
prob_teo_b <- c(P = pnorm(dP_b), Q = pnorm(d2_b), QV = pnorm(d1_b))

# IC de Monte Carlo al 95 %: p estimada ± 1.96 sqrt[p(1-p)/N].
margen_b <- qnorm(0.975) * sqrt(prob_mc_b * (1 - prob_mc_b) / n_b)
tabla_prob_b <- data.frame(
  Medida = names(prob_mc_b),
  MC = 100 * prob_mc_b,
  Teorica = 100 * prob_teo_b,
  Inferior = 100 * (prob_mc_b - margen_b),
  Superior = 100 * (prob_mc_b + margen_b),
  row.names = NULL
)
cat("\nProbabilidades en % | MC: Monte Carlo | IC al 95 %\n")
print(tabla_prob_b, digits = 6, row.names = FALSE)

# 10. Reconstruir el precio de la call y contrastar con (a) ----
# Valor de recibir la acción y de pagar K, ambos condicionados al evento A.
activo_b <- S0 * exp(-q * T) * prob_mc_b[["QV"]]
efectivo_b <- K * exp(-r * T) * prob_mc_b[["Q"]]
precio_mc_b <- activo_b - efectivo_b
precio_teo_b <- S0 * exp(-q * T) * pnorm(d1_b) - K * exp(-r * T) * pnorm(d2_b)

tabla_precio_b <- data.frame(
  Concepto = c("Entrega de la acción (MC)", "Pago del strike (MC)",
               "Call reconstruida (MC)", "Call con probabilidades teóricas",
               "Black-Scholes de (a)"),
  USD = c(activo_b, efectivo_b, precio_mc_b, precio_teo_b, referencia)
)
cat("\nDescomposición del precio, en USD por acción\n")
print(tabla_precio_b, digits = 6, row.names = FALSE)
# Las dos expresiones teóricas deben coincidir salvo redondeo numérico.
stopifnot(isTRUE(all.equal(precio_teo_b, referencia, tolerance = 1e-10)))


# 11. Funciones comunes para las partes (c), (d) y (e) ----
# En (a)-(b) usamos sigma histórica; aquí usamos la IV de pantalla de la call.
q_referencia <- q
K_principal <- K
sigma_principal <- call_principal$iv
stopifnot(T > 0, S0 > 0, is.finite(r), sigma_principal > 0)

# Black-Scholes europeo con dividendo continuo; admite varios strikes o plazos.
bs_q <- function(S, K, r, sigma, tau, q = 0, tipo = c("call", "put")) {
  tipo <- match.arg(tipo)
  n <- max(length(S), length(K), length(tau))
  S <- rep_len(S, n); K <- rep_len(K, n); tau <- rep_len(tau, n)
  Sq <- S * exp(-q * pmax(tau, 0))
  fwd <- Sq - K * exp(-r * pmax(tau, 0))
  out <- if (tipo == "call") pmax(fwd, 0) else pmax(-fwd, 0)
  vivo <- tau > 0 & sigma > 0
  if (any(vivo)) {
    s <- Sq[vivo]; k <- K[vivo]; tt <- tau[vivo]
    d1 <- (log(s / k) + (r + sigma^2 / 2) * tt) / (sigma * sqrt(tt))
    d2 <- d1 - sigma * sqrt(tt)
    out[vivo] <- if (tipo == "call") {
      s * pnorm(d1) - k * exp(-r * tt) * pnorm(d2)
    } else {
      k * exp(-r * tt) * pnorm(-d2) - s * pnorm(-d1)
    }
  }
  out
}

# Black-76 descuenta el pago expresado en términos del forward.
black76 <- function(F, K, r, sigma, tau, tipo = "call") {
  bs_q(F, K, r, sigma, tau, q = r, tipo = tipo)
}

# Despeje de q a partir de C-P bajo paridad europea.
q_implicito_ntr <- function(cp, K, S, tasa, plazo) {
  argumento <- (cp + K * exp(-tasa * plazo)) / S

  resultado <- ifelse(argumento > 0,-log(argumento) / plazo,NA_real_)

  return(resultado)
}

# 12. Verificación con el caso del curso ----
# S = K = 100, r = 5 %, sigma = 20 %, T = 1 y q = 3 %.
caso_c <- bs_q(100, 100, 0.05, 0.20, 1, 0.03, "call")
caso_p <- bs_q(100, 100, 0.05, 0.20, 1, 0.03, "put")
caso_d1 <- (0.05 - 0.03 + 0.5 * 0.20^2) / 0.20
caso_f <- 100 * exp(0.05 - 0.03)
caso_q <- q_implicito_ntr(caso_c - caso_p, 100, 100, 0.05, 1)
caso_black76 <- black76(caso_f, 100, 0.05, 0.20, 1)
stopifnot(
  abs(caso_c - 8.6525) < 1e-4,
  abs(pnorm(caso_d1) - 0.5793) < 1e-4,
  abs(caso_f - 102.02) < 0.005,
  abs(caso_q - 0.03) < 1e-10,
  abs(caso_c - caso_black76) < 1e-10
)
cat("\nVerificación del caso del curso: OK\n")
print(data.frame(Call_BS = caso_c, Probabilidad_Nd1 = pnorm(caso_d1),
                 Forward = caso_f, Dividendo_implicito = caso_q,
                 Call_Black76 = caso_black76), row.names = FALSE)

# PARTE (c). Dividendo implícito, paridad y precio de la call ----

# 13.1. Plazo, tasa y precio de la bitacora 1 ----
dias_opcion <- as.integer(fecha_vencimiento - fecha_valoracion)   # 46
T <- dias_opcion / 365

stopifnot(T > 0, S0 > 0, is.finite(r), sigma_principal > 0)
cat("\nr (continua, base 365):", round(100 * r, 4), "% | T =", round(T, 5), "\n")

# Comparación opcional: falta un precio comparable documentado de la bitácora 1.
precio_bitacora1       <- NA_real_   # <- precio de tu bitacora 1
bitacora1_misma_opcion <- NA         # TRUE si mismo strike y vencimiento
bitacora1_descripcion  <- "Completar: fecha, S0, K, sigma, r, metodo de la bitacora 1"

# 13.2. Dividendos declarados vs q de referencia ----
dividendos <- data.frame(
  fecha_ex   = as.Date(c("2025-12-31", "2026-09-29")),
  fecha_pago = as.Date(c("2026-01-16", "2026-10-16")),
  monto      = c(0.545, 0.55),
  fuente     = c("Nutrien/Business Wire 2025-11-04", "dividendpedia.com (captura)"),
  stringsAsFactors = FALSE
)
ex_proximo_estimado <- as.Date("2026-12-31")   # ESTIMADA, no declarada

dividendos_periodo <- subset(dividendos,
                             fecha_ex > fecha_valoracion & fecha_ex <= fecha_vencimiento)

if (nrow(dividendos_periodo) == 0) {
  VP_dividendos <- 0
} else {
  t_div <- as.numeric(dividendos_periodo$fecha_ex - fecha_valoracion) / 365
  VP_dividendos <- sum(dividendos_periodo$monto * exp(-r * t_div))
}
stopifnot(is.finite(VP_dividendos), VP_dividendos < S0)
# q continuo equivalente al VP de los dividendos registrados en el plazo
q_referencia <- -log((S0 - VP_dividendos) / S0) / T

cat("\n(c) CALENDARIO DE DIVIDENDOS REGISTRADO EN EL RMD\n")
print(dividendos, row.names = FALSE)
cat("Ex-dividendos declarados dentro del plazo:", nrow(dividendos_periodo),
    "| q de referencia:", round(100 * q_referencia, 4), "%\n")
cat("Proximo ex-dividendo (estimado):", format(ex_proximo_estimado), "\n")

# 13.3. Dividendo implicito por strike ----
# C - P = S0 exp(-qT) - K exp(-rT)  =>  q(K) = -log[(C-P+K exp(-rT))/S0]/T
C <- cadena_opciones

C$cp_mid  <- C$call_mid - C$put_mid
C$cp_bajo <- C$call_bid - C$put_ask     # extremo inferior de C-P
C$cp_alto <- C$call_ask - C$put_bid     # extremo superior de C-P

C$q_mid <- q_implicito_ntr(C$cp_mid,  C$strike, S0, r, T)
C$q_min <- q_implicito_ntr(C$cp_alto, C$strike, S0, r, T)   # C-P alto => q bajo
C$q_max <- q_implicito_ntr(C$cp_bajo, C$strike, S0, r, T)

C$q_mid_pct <- 100 * C$q_mid
C$q_min_pct <- 100 * C$q_min
C$q_max_pct <- 100 * C$q_max
C$q_incluye_referencia <- with(C, q_min <= q_referencia & q_max >= q_referencia)

# 13.4. Ejercicio temprano de las puts (la paridad es EUROPEA) ----
# Las calls sin dividendos en el plazo no se ejercen antes; las puts si
# pueden. Se clasifican por delta de la put (criterio heuristico).
C$grupo <- cut(C$put_delta, breaks = c(-Inf, -0.8, -0.5, Inf),
               labels = c("put muy ITM", "put ITM", "put OTM/ATM"))

# 13.5. Resumen por muestra y chequeo falsable ----
resumen_muestra <- function(idx, nombre) {
  data.frame(
    muestra = nombre, n = sum(idx),
    q_mid_promedio_pct = mean(C$q_mid_pct[idx], na.rm = TRUE),
    q_mid_min_pct = min(C$q_mid_pct[idx], na.rm = TRUE),
    q_mid_max_pct = max(C$q_mid_pct[idx], na.rm = TRUE),
    interseccion_bandas_inf_pct = max(C$q_min_pct[idx], na.rm = TRUE),
    interseccion_bandas_sup_pct = min(C$q_max_pct[idx], na.rm = TRUE)
  )
}
resumen_q <- rbind(
  resumen_muestra(rep(TRUE, nrow(C)), "todas"),
  resumen_muestra(C$grupo == "put OTM/ATM", "solo puts OTM/ATM")
)
# Si hubiera un unico q para todos los strikes, las bandas bid/ask deben
# tener interseccion no vacia. Este chequeo PUEDE fallar.
resumen_q$interseccion_vacia <-
  resumen_q$interseccion_bandas_inf_pct > resumen_q$interseccion_bandas_sup_pct
resumen_q$q_ref_en_interseccion <-
  100 * q_referencia >= resumen_q$interseccion_bandas_inf_pct &
  100 * q_referencia <= resumen_q$interseccion_bandas_sup_pct

# 13.6. Cotas de las cotizaciones ----
# Un bid bajo la cota es un diagnóstico; no prueba arbitraje ejecutable.
C$call_bid_bajo_cota <- C$call_bid <
  pmax(S0 * exp(-q_referencia * T) - C$strike * exp(-r * T), 0)
C$put_bid_bajo_intrinseco <- C$put_bid < pmax(C$strike - S0, 0)  # put americana

cat("\n(c) DIVIDENDOS IMPLICITOS POR STRIKE\n")
print(C[c("strike", "grupo", "q_mid_pct", "q_min_pct", "q_max_pct",
          "q_incluye_referencia", "call_bid_bajo_cota", "put_bid_bajo_intrinseco")],
      digits = 4, row.names = FALSE)
cat("\nRESUMEN POR MUESTRA (q anual, %)\n")
print(resumen_q, digits = 4, row.names = FALSE)
cat("Bandas que contienen q de referencia:", sum(C$q_incluye_referencia, na.rm = TRUE),
    "de", nrow(C), "\n")

# 13.7. Precio de la call principal (Prop. 16.5 / 16.9) ----
p <- C[C$strike == K_principal, , drop = FALSE]
stopifnot(nrow(p) == 1L)

# La IV propia proviene de la cotización; su cercanía al mercado no es validación independiente.
# Prueba menos circular: IV de los strikes vecinos (sin el K principal).
vecinos <- which(C$strike %in% c(67.5, 72.5))
sigma_vecinos <- mean(C$call_iv_pct[vecinos]) / 100
# Reutiliza la volatilidad histórica estimada con el Excel en (a).
sigma_historica <- sigma

precio_call <- function(sigma, q) {
  as.numeric(bs_q(S0, K_principal, r, sigma, T, q, "call"))
}
comparacion <- data.frame(
  modelo = c("BS, IV propia, q = q ref.",
             "BS, IV propia, q = q implicito (Prop. 16.9)",
             "BS, IV vecinos (67.5 y 72.5), q = q ref."),
  precio = c(precio_call(sigma_principal, q_referencia),
             precio_call(sigma_principal, p$q_mid),
             precio_call(sigma_vecinos, q_referencia))
)
if (is.finite(sigma_historica)) {
  comparacion <- rbind(comparacion,
    data.frame(modelo = "BS, sigma historica, q = q ref.",
               precio = precio_call(sigma_historica, q_referencia)))
}
comparacion$vs_mid <- comparacion$precio - p$call_mid
comparacion$dentro_bid_ask <- comparacion$precio >= p$call_bid &
                              comparacion$precio <= p$call_ask

cat("\nCALL K =", K_principal, "| mercado: bid", p$call_bid, "mid", p$call_mid,
    "ask", p$call_ask, "\n")
print(comparacion, digits = 5, row.names = FALSE)

if (is.finite(precio_bitacora1)) {
  cat("Precio bitacora 1:", round(precio_bitacora1, 4),
      "| misma opcion:", bitacora1_misma_opcion, "\n",
      bitacora1_descripcion, "\n",
      "Diferencia BS(q ref.) - bitacora 1:",
      round(comparacion$precio[1] - precio_bitacora1, 4), "USD\n")
} else {
  cat("Precio bitacora 1: PENDIENTE (rellenar precio_bitacora1).\n")
}

C$call_ask_bajo_cota <- C$call_ask <
     pmax(S0 * exp(-q_referencia * T) - C$strike * exp(-r * T), 0)
C$put_ask_bajo_intrinseco <- C$put_ask < pmax(C$strike - S0, 0)
cat("Calls con ask bajo la cota:", sum(C$call_ask_bajo_cota),
       "| Puts con ask bajo el valor intrinseco:", sum(C$put_ask_bajo_intrinseco), "\n")
# 13.8. Verificaciones ----
stopifnot(
  all(C$q_min <= C$q_max, na.rm = TRUE),
  # ida y vuelta: q implicito reproduce C-P (prueba la implementacion)
  all(abs(S0 * exp(-C$q_mid * T) - C$strike * exp(-r * T) - C$cp_mid) < 1e-9,
      na.rm = TRUE)
)
cat("\nVERIFICACIONES DE (c): OK\n")
cat("Calls con bid bajo la cota de no arbitraje:", sum(C$call_bid_bajo_cota),
    "| Puts con bid bajo el valor intrinseco:", sum(C$put_bid_bajo_intrinseco), "\n")
cat("ADVERTENCIA: paridad europea; las puts son americanas.",
    "La r es de la letra de 4 semanas, no observada a 46 dias.\n")

# PARTE (d). Forward implícito y valoración con Black-76 ----

# Paridad:
# F(K) = K + exp(r*T)*(C-P)

# Forward usando precios medios
C$f_mid <- C$strike + exp(r * T) * C$cp_mid

# Banda inferior
C$f_min <- C$strike + exp(r * T) * C$cp_bajo

# Banda superior
C$f_max <- C$strike + exp(r * T) * C$cp_alto

# Forward teorico por costo de acarreo
F_carry <- S0 * exp((r - q_referencia) * T)

# Verificar si el forward teorico cae en la banda
C$f_incluye_carry <- with(C,f_min <= F_carry &f_max >= F_carry)

# Residuo respecto a paridad europea de referencia
C$residuo_paridad <- C$cp_mid -
  (S0 * exp(-q_referencia * T) -C$strike * exp(-r * T))

cat("\n(d) FORWARDS IMPLICITOS\n")

print(
  C[
    c(
      "strike",
      "f_mid",
      "f_min",
      "f_max",
      "f_incluye_carry"
    )
  ],
  digits = 6,
  row.names = FALSE
)

# 14.1. Friccion y cotizaciones (se calculan ANTES de usarlas) ----
C$spread_par <- C$call_spread + C$put_spread
C$ancho_forward_ejecutable <- C$f_max - C$f_min
C$last_call_fuera <- C$call_last < C$call_bid | C$call_last > C$call_ask
C$last_put_fuera  <- C$put_last  < C$put_bid  | C$put_last  > C$put_ask

# 14.2. Forward implicito: es el mismo para todos los strikes? ----
f_inf <- max(C$f_min)     # con un unico forward, las bandas se intersectan
f_sup <- min(C$f_max)
resumen_f <- data.frame(
  f_mid_promedio = mean(C$f_mid),
  f_mid_sd       = sd(C$f_mid),
  f_mid_min      = min(C$f_mid),
  f_mid_max      = max(C$f_mid),
  interseccion_inf = f_inf,
  interseccion_sup = f_sup,
  F_carry        = F_carry
)
resumen_f$interseccion_vacia    <- f_inf > f_sup            # puede fallar
resumen_f$carry_en_interseccion <- F_carry >= f_inf & F_carry <= f_sup
cat("\n(d) RESUMEN DEL FORWARD IMPLICITO\n")
print(resumen_f, digits = 6, row.names = FALSE)

# 14.3. Valoracion de la call principal ----
p <- C[C$strike == K_principal, , drop = FALSE]
stopifnot(nrow(p) == 1L)

precio_BS_q0     <- as.numeric(bs_q(S0, K_principal, r, sigma_principal, T, q_referencia, "call"))
precio_BS_qimpl  <- as.numeric(bs_q(S0, K_principal, r, sigma_principal, T, p$q_mid, "call"))
precio_B76_carry <- as.numeric(black76(F_carry,  K_principal, r, sigma_principal, T, "call"))
precio_B76_impl  <- as.numeric(black76(p$f_mid, K_principal, r, sigma_principal, T, "call"))

# 14.4. Verificaciones ----
# Black-76 con f implicito = BS con q implicito: f_impl = K + exp(rT)(C-P)
# = S0 exp((r-q_impl)T), y Black-76 con F = S0 exp((r-q)T) es la misma
# formula que BS con dividendo q. Identidad algebraica (prueba la
# implementacion, no el mercado).
stopifnot(
  length(C$spread_par) == nrow(C),
  abs(p$f_mid - S0 * exp((r - p$q_mid) * T)) < 1e-9,
  abs(precio_BS_q0 - precio_B76_carry) < 1e-9,
  abs(precio_BS_qimpl - precio_B76_impl) < 1e-9,
  all(C$q_min <= C$q_max, na.rm = TRUE),
  all(C$f_min <= C$f_max),
  all(abs(C$ancho_forward_ejecutable - exp(r * T) * C$spread_par) < 1e-8)
)
cat("\nVERIFICACIONES ALGEBRAICAS: OK\n")

# 14.5. Resultados de la call principal ----
cat("\nRESULTADOS DE LA CALL PRINCIPAL\n")
cat("Strike:", K_principal, "| S0:", S0, "| sigma:", sigma_principal, "\n")
cat("Mercado bid:", p$call_bid, "| mid:", p$call_mid, "| ask:", p$call_ask, "\n")
cat("Dividendo implicito anual (%):", round(100 * p$q_mid, 6), "\n")
cat("Forward por costo de acarreo:", round(F_carry, 6),
    "| forward implicito:", round(p$f_mid, 6), "\n")
cat("BS con q referencia:", round(precio_BS_q0, 6),
    "| BS con q implicito:", round(precio_BS_qimpl, 6), "\n")
cat("Black-76 carry:", round(precio_B76_carry, 6),
    "| Black-76 con f implicito:", round(precio_B76_impl, 6), "\n")
cat("BS(q referencia) - mid:", round(precio_BS_q0 - p$call_mid, 6), "USD\n")
cat("Spread call:", p$call_spread, "| spread put:", p$put_spread,
    "| combinado:", p$spread_par, "USD\n")

cat("\nRESUMEN DE LA CADENA\n")
cat("Bandas de q que contienen q referencia:", sum(C$q_incluye_referencia, na.rm = TRUE),
    "de", nrow(C), "\n")
cat("Bandas de forward que contienen carry:", sum(C$f_incluye_carry, na.rm = TRUE),
    "de", nrow(C), "\n")
cat("Ancho de banda del forward (min, max):", round(range(C$ancho_forward_ejecutable), 3), "\n")
cat("Calls con last fuera de bid/ask:", sum(C$last_call_fuera),
    "| Puts con last fuera de bid/ask:", sum(C$last_put_fuera), "\n")
cat("\nADVERTENCIA: la paridad utilizada es europea; las puts son americanas.\n")

# PARTE (e). Dividendo continuo frente a dividendos discretos ----
# 15.1. Reutilizar el calendario y el VP calculados en (c) ----
cat("\n(e) Dividendos discretos, continuos y escenario hipotético\n")
precio_q0 <- as.numeric(bs_q(S0, K_principal, r, sigma_principal, T, 0, "call"))
precio_discreto <- as.numeric(bs_q(S0 - VP_dividendos, K_principal, r,
                                   sigma_principal, T, 0, "call"))

# 15.2. Lo que supondria el modelo continuo ----
# Rendimiento anual indicativo: 4 dividendos de 0.55 sobre S0.
q_anual <- 4 * 0.55 / S0

precio_continuo <- as.numeric(bs_q(S0, K_principal, r, sigma_principal, T,
                                   q_anual, "call"))

costo_usd <- precio_q0 - precio_continuo          # sobrevaloracion del dividendo
costo_pct <- 100 * costo_usd / precio_q0
medio_spread <- p$call_spread / 2

# Forward que implicaria el supuesto continuo vs el real (sin dividendo)
F_continuo <- S0 * exp((r - q_anual) * T)

# 15.3. Escenario hipotetico: ex-dividendo dentro del plazo ----
# NO es un dividendo anunciado. Es una sensibilidad.
fecha_ex_hip <- as.Date("2026-12-01")
t_hip <- as.numeric(fecha_ex_hip - fecha_valoracion) / 365
stopifnot(t_hip > 0, t_hip < T)
VP_hip <- 0.55 * exp(-r * t_hip)

# (i) ajuste del spot por VP (aproximacion habitual)
precio_hip_vp <- as.numeric(bs_q(S0 - VP_hip, K_principal, r, sigma_principal, T, 0, "call"))
# Nota: el dividendo proporcional de Bjork Prop. 16.5 con delta = VP/S0
# da exactamente el mismo precio (S0*(1-delta) = S0 - VP). Es algebraico,
# no un chequeo independiente.

efecto_hip <- precio_q0 - precio_hip_vp

# 15.4. Tabla y chequeos ----
resultados_e <- data.frame(
  escenario = c("Calendario registrado: sin ex-dividendo (q = 0)",
                "Dividendos registrados: ajuste por VP",
                "Supuesto continuo, q anual indicativo",
                "Hipotetico: 0.55 el 1-dic (ajuste VP)"),
  precio_call = c(precio_q0, precio_discreto, precio_continuo,
                  precio_hip_vp)
)
resultados_e$dentro_bid_ask <- resultados_e$precio_call >= p$call_bid &
                               resultados_e$precio_call <= p$call_ask
cat("\nRESULTADOS DEL APARTADO (e), call K =", K_principal, "\n")
print(resultados_e, row.names = FALSE, digits = 6)

cat("\nq anual indicativo (%):", round(100 * q_anual, 4), "\n")
cat("Costo de imponer el supuesto continuo:", round(costo_usd, 4), "USD (",
    round(costo_pct, 2), "% del precio ) =",
    round(costo_usd / medio_spread, 2), "x medio spread (", medio_spread, ")\n")
cat("Costo del dividendo hipotetico:", round(efecto_hip, 4), "USD =",
    round(efecto_hip / medio_spread, 2), "x medio spread\n")
cat("Forward con q anual:", round(F_continuo, 4),
    "| con q = 0:", round(F_carry, 4), "\n")

# Limitacion declarada: la cadena sola no discrimina entre q = 0 y q anual.
q_en_banda  <- sum(C$q_min <= q_anual & C$q_max >= q_anual, na.rm = TRUE)
f_en_banda  <- sum(C$f_min <= F_continuo & C$f_max >= F_continuo)
cat("Bandas de q que contienen q anual:", q_en_banda, "de", nrow(C), "\n")
cat("Bandas de forward que contienen F con q anual:", f_en_banda, "de", nrow(C), "\n")

stopifnot(
  abs(precio_discreto - precio_call(sigma_principal, q_referencia)) < 1e-9,
  precio_continuo < precio_q0,
  precio_hip_vp < precio_q0
)
cat("\nVerificaciones de (e): OK\n")
cat("El calendario registrado tiene", nrow(dividendos_periodo),
    "ex-dividendos dentro del plazo. La próxima fecha estimada es",
    format(ex_proximo_estimado), "(no se trata como anuncio confirmado).\n")
cat("El q anual indicativo reduce la call en", round(costo_usd, 4),
    "USD; el escenario de diciembre es hipotético.\n")

cat("\nBitácora 2: análisis (a)-(e) finalizado. Resultados mostrados en consola.\n")
