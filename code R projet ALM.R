################################################################################
##                                                                            ##
##   EXAMEN FINAL — INGÉNIERIE FINANCIÈRE & GESTION OBLIGATAIRE               ##
##   Modélisation ALM | Arbitrage de Courbe | Conformité COSUMAF              ##
##             4ème Année — 2025-2026                                         ##
##                                                                            ##
################################################################################
##
##  STRUCTURE :
##  ─────────────────────────────────────────────────────────────────────────
##  ÉTAPE 0  : Installation des packages & Cartographie des données
##  ÉTAPE 1  : Collecte, préparation et lissage de la courbe des taux
##             (Moyennes Mobiles vs Lissage de Holt + Analyse des résidus)
##  ÉTAPE 2  : Économétrie univariée — Modèles GARCH par pays/segment
##             (Tests ADF, ARCH, Jarque-Bera + ARMA-GARCH sur les flux passif)
##  ÉTAPE 3  : Économétrie multivariée — Modèles VAR
##             (VAR National Cameroun + VAR Régional + IRF + FEVD)
##  ÉTAPE 4  : Évaluation du risque ALM
##             (Liquidity at Risk, Gap de liquidité, Performance par pays)
##  ÉTAPE 5  : Stress Testing & Conformité COSUMAF
##             (Double Choc Actif/Passif + Vérification des ratios)
##  ÉTAPE 6  : Export des résultats pour l'outil VBA
##             
################################################################################


################################################################################
## ÉTAPE 0 — INSTALLATION DES PACKAGES & CHARGEMENT DES DONNÉES
################################################################################

# ─── 0.1 Installation automatique si besoin ────────────────────────────────

packages_requis <- c(
  "readxl",      # Import des fichiers Excel
  "dplyr",       # Manipulation de données
  "tidyr",       # Restructuration de tableaux
  "lubridate",   # Gestion des dates
  "zoo",         # Moyennes mobiles (rollmean)
  "forecast",    # Lissage de Holt, auto.arima
  "tseries",     # Test ADF (adf.test), Jarque-Bera
  "FinTS",       # Test ARCH (ArchTest)
  "rugarch",     # Estimation GARCH (ugarchspec, ugarchfit)
  "vars",        # Modèles VAR, IRF, FEVD
  "urca",        # Test de Johansen (cointégration)
  "ggplot2",     # Graphiques professionnels
  "gridExtra",   # Disposition de plusieurs graphiques
  "scales",      # Formatage des axes ggplot
  "writexl"      # Export Excel des résultats
)

nouveaux <- packages_requis[!packages_requis %in% installed.packages()[, "Package"]]
if (length(nouveaux) > 0) {
  cat(">>> Installation de", length(nouveaux), "packages manquants...\n")
  install.packages(nouveaux, dependencies = TRUE)
}

suppressPackageStartupMessages({
  lapply(packages_requis, library, character.only = TRUE)
})
cat("✔ Tous les packages sont chargés.\n\n")

# ─── 0.2 Chargement de la Data Room ────────────────────────────────────────

# Adapter ces chemins si vos fichiers sont dans un sous-dossier
FICHIER_ACTIF_PASSIF <- "SUIVI_PORTEFEUILLES_ACTIF PASSIF_2025.xlsx"
FICHIER_TMP          <- "TMP_COMPLET_FINAL.xlsx"

# Lecture de la feuille ACTIF (18 lignes obligataires)
actif_raw <- readxl::read_excel(FICHIER_ACTIF_PASSIF, sheet = "ACTIF",
                                skip = 20, col_names = FALSE)
# Nettoyage : retrait des lignes vides et attribution de noms de colonnes
actif_raw <- actif_raw[!is.na(actif_raw[[1]]), ]
noms_actif <- c("Libelle", "Emetteur", "Date_operation", "Taux_facial",
                "Nb_titres", "Date_dernier_coupon", "Nominal_brut",
                "Coupon_couru", "Valorisation")
# On sélectionne uniquement les 18 lignes de titres (avant les lignes TOTAL)
actif <- actif_raw[1:18, 1:9]


colnames(actif) <- noms_actif

# Conversion des types
actif$Taux_facial    <- as.numeric(actif$Taux_facial)
actif$Nb_titres      <- as.numeric(actif$Nb_titres)
actif$Nominal_brut   <- as.numeric(actif$Nominal_brut)
actif$Coupon_couru   <- as.numeric(actif$Coupon_couru)
actif$Valorisation   <- as.numeric(actif$Valorisation)

# Montant espèces
ESPECES_FCFA <- 865222220

# Lecture de la feuille PASSIF — extraction des totaux de parts par semaine
# On reconstruit manuellement depuis les données brutes observées
# (les totaux cumulés de fin de chaque semaine de VL)
semaines_vl <- as.Date(c("2026-01-02","2026-01-09","2026-01-16","2026-01-23",
                         "2026-01-30","2026-02-06","2026-02-13","2026-02-20",
                         "2026-02-27","2026-03-06","2026-03-13","2026-03-20",
                         "2026-03-27","2026-04-03","2026-04-19","2026-04-17",
                         "2026-04-24","2026-05-01","2026-05-08"))

parts_cumul <- c(5248.06, 7941.63, 7993.70, 8331.46, 8437.01, 8551.30,
                 8696.47, 10838.00, 10933.80, 11063.40, 13367.96,
                 13471.27, 13637.64, 13977.60, 14359.64, 14519.48,
                 14769.21,14930.60,15238.97)

passif_df <- data.frame(
  Semaine     = semaines_vl,
  Parts_total = parts_cumul,
  stringsAsFactors = FALSE
)
passif_df$Flux_net <- c(NA, diff(passif_df$Parts_total))

# Lecture du fichier TMP — feuille FORMAT_LONG
tmp_long <- readxl::read_excel(FICHIER_TMP, sheet = "FORMAT_LONG", skip = 1)
colnames(tmp_long) <- c("Date_adj", "Pays", "Instrument", "Maturite",
                        "TMP_pct", "Trimestre")

# Conversion des dates Excel (numérique -> Date R)
tmp_long$Date <- as.Date(as.numeric(tmp_long$Date_adj), origin = "1899-12-30")
tmp_long$TMP  <- as.numeric(tmp_long$TMP_pct)
tmp_long      <- tmp_long[!is.na(tmp_long$TMP) & !is.na(tmp_long$Date), ]

cat("✔ Données chargées :\n")
cat("  - Actif          :", nrow(actif), "lignes obligataires\n")
cat("  - Passif         :", nrow(passif_df), "semaines de VL\n")
cat("  - TMP (marché)   :", nrow(tmp_long), "observations\n\n")
# ─── 0.3 Cartographie rapide du portefeuille ────────────────────────────────

cat("═══════════════════════════════════════════════════════\n")
cat("  CARTOGRAPHIE DU PORTEFEUILLE ACTIF\n")
cat("═══════════════════════════════════════════════════════\n")

actif$Emetteur_clean <- trimws(actif$Emetteur)

# Regroupement par émetteur
resume_emetteur <- actif %>%
  group_by(Emetteur_clean) %>%
  summarise(
    Nb_lignes       = n(),
    Valorisation    = sum(Valorisation, na.rm = TRUE),
    Coupon_annuel   = sum(Nominal_brut * Taux_facial, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  mutate(
    Poids_pct = Valorisation / sum(Valorisation) * 100
  ) %>%
  arrange(desc(Valorisation))

total_oblig  <- sum(actif$Valorisation, na.rm = TRUE)
actif_brut   <- total_oblig + ESPECES_FCFA

cat(sprintf("  Total obligations  : %s FCFA\n",
            format(round(total_oblig), big.mark = " ")))
cat(sprintf("  Espèces            : %s FCFA\n",
            format(ESPECES_FCFA, big.mark = " ")))
cat(sprintf("  Actif brut estimé  : %s FCFA\n\n",
            format(round(actif_brut), big.mark = " ")))

cat("  Répartition par émetteur :\n")
print(resume_emetteur, digits = 2)

# Alerte COSUMAF immédiate
limite_cosumaf <- 35
cat("\n  Vérification COSUMAF (limite 35% par émetteur) :\n")
for (i in seq_len(nrow(resume_emetteur))) {
  poids <- resume_emetteur$Poids_pct[i]
  em    <- resume_emetteur$Emetteur_clean[i]
  flag  <- if (poids > limite_cosumaf) "⚠️  ALERTE" else "✔ Conforme"
  cat(sprintf("    %-30s : %5.1f%%  [%s]\n", em, poids, flag))
}

VL_PAR_PART_APPROX <- actif_brut / max(parts_cumul)
cat(sprintf("\n  VL par part (approx.) : %.2f FCFA\n", VL_PAR_PART_APPROX))
cat(sprintf("  Parts en circulation  : %.0f\n\n", max(parts_cumul)))




################################################################################
## ÉTAPE 1 — LISSAGE DE LA COURBE DES TAUX
##
##  Objectif : sur un marché d'adjudications (données irrégulières, peu
##  liquides), distinguer le signal de taux du bruit d'illiquidité.
##  Deux méthodes comparées :
##    (A) Moyennes Mobiles (MM3 et MM5)
##    (B) Lissage exponentiel de Holt (double lissage avec tendance)
################################################################################

cat("═══════════════════════════════════════════════════════\n")
cat("  ÉTAPE 1 — LISSAGE DE LA COURBE DES TAUX\n")
cat("═══════════════════════════════════════════════════════\n\n")


# ─── 1.1 Fonction utilitaire : construction d'une série temporelle par segment

construire_serie <- function(pays, instrument, maturite, data = tmp_long) {
  # Filtre et tri chronologique
  sous <- data %>%
    filter(Pays == pays, Instrument == instrument, Maturite == maturite) %>%
    arrange(Date) %>%
    dplyr::select(Date, TMP) # Explicitly use dplyr::select
  
  if (nrow(sous) < 4) {
    cat(sprintf("  [ATTENTION] Série %s %s %s : seulement %d observations.\n",
                pays, instrument, maturite, nrow(sous)))
  }
  return(sous)
}

# ─── 1.2 Lissage MM et Holt — fonction générique

lisser_serie <- function(serie_df, label = "", k_mm = c(3, 5)) {
  # serie_df : data.frame avec colonnes Date et TMP
  # label    : étiquette pour les graphiques
  
  y <- serie_df$TMP
  n <- length(y)
  
  if (n < 4) {
    cat(sprintf("  [SKIP] %s : pas assez de données (%d obs).\n", label, n))
    return(NULL)
  }
  
  # --- Méthode A : Moyennes Mobiles ---
  mm3 <- as.numeric(rollmean(y, k = k_mm[1], align = "center", fill = NA))
  mm5 <- if (n >= k_mm[2]) {
    as.numeric(rollmean(y, k = k_mm[2], align = "center", fill = NA))
  } else {
    rep(NA, n)
  }
  
  # --- Méthode B : Lissage de Holt ---
  ts_y     <- ts(y, frequency = 1)
  holt_fit <- tryCatch(
    holt(ts_y, h = 4, damped = FALSE),
    error = function(e) NULL
  )
  
  if (is.null(holt_fit)) {
    cat(sprintf("  [ATTENTION] Holt a échoué pour %s. Utilisation de SES.\n",
                label))
    holt_fit <- ses(ts_y, h = 4)
  }
  
  holt_lisse <- as.numeric(fitted(holt_fit))
  
  # --- Résidus ---
  residus_mm3  <- y - mm3
  residus_holt <- y - holt_lisse
  
  # --- Métriques de comparaison ---
  rmse_mm3  <- sqrt(mean(residus_mm3^2,  na.rm = TRUE))
  rmse_holt <- sqrt(mean(residus_holt^2, na.rm = TRUE))
  mae_mm3   <- mean(abs(residus_mm3),    na.rm = TRUE)
mae_holt  <- mean(abs(residus_holt),   na.rm = TRUE)

# --- Test Ljung-Box (H0 : résidus = bruit blanc) ---
lag_test  <- min(10, max(4, floor(n / 3)))
lb_mm3    <- tryCatch(Box.test(na.omit(residus_mm3),
                               lag = lag_test, type = "Ljung-Box"),
                      error = function(e) list(p.value = NA))
lb_holt   <- tryCatch(Box.test(na.omit(residus_holt),
                               lag = lag_test, type = "Ljung-Box"),
                      error = function(e) list(p.value = NA))

cat(sprintf("\n  ── Série : %s  (n = %d) ──\n", label, n))
cat(sprintf("     MM3   — RMSE : %.4f  |  MAE : %.4f  |  Ljung-Box p = %.3f\n",
            rmse_mm3, mae_mm3, lb_mm3$p.value))
cat(sprintf("     Holt  — RMSE : %.4f  |  MAE : %.4f  |  Ljung-Box p = %.3f\n",
            rmse_holt, mae_holt, lb_holt$p.value))

meilleure <- if (rmse_holt < rmse_mm3) "Holt" else "MM3"
cat(sprintf("     → Meilleure méthode (RMSE) : %s\n", meilleure))

# Interprétation du Ljung-Box
if (!is.na(lb_holt$p.value) && lb_holt$p.value > 0.05) {
  cat("     → Holt : résidus non autocorrélés (p > 0.05) ✔ bruit blanc\n")
} else {
  cat("     → Holt : structure résiduelle encore présente (p <= 0.05)\n")
}

# Extraction des paramètres Holt
alpha_opt <- round(holt_fit$model$par["alpha"], 4)
beta_opt  <- round(holt_fit$model$par["beta"],  4)
cat(sprintf("     → Holt alpha = %.4f  |  beta = %.4f\n",
            alpha_opt, beta_opt))

# Prévisions 4 semaines
prev4     <- as.numeric(holt_fit$mean)
cat(sprintf("     → Prévisions (4 pas) Holt : %.3f  %.3f  %.3f  %.3f\n",
            prev4[1], prev4[2], prev4[3], prev4[4]))

# Graphique
df_plot <- data.frame(
  Date  = serie_df$Date,
  Brut  = y,
  MM3   = mm3,
  MM5   = mm5,
  Holt  = holt_lisse
)

g <- ggplot(df_plot, aes(x = Date)) +
  geom_line(aes(y = Brut, colour = "Brut"),   linewidth = 0.8,  linetype = "dotted") + # Changed size to linewidth
  geom_line(aes(y = MM3,  colour = "MM3"),    linewidth = 1.0,  na.rm = TRUE) + # Changed size to linewidth
  geom_line(aes(y = MM5,  colour = "MM5"),    linewidth = 1.0,  na.rm = TRUE) + # Changed size to linewidth
  geom_line(aes(y = Holt, colour = "Holt"),   linewidth = 1.2,  na.rm = TRUE) + # Changed size to linewidth
  scale_colour_manual(
    values = c("Brut" = "grey60", "MM3" = "steelblue",
               "MM5" = "tomato",  "Holt" = "darkgreen"),
    name = "Méthode"
  ) +
  labs(title   = paste("Lissage de la courbe —", label),
       subtitle = sprintf("MM3 RMSE=%.4f | Holt RMSE=%.4f | Meilleure: %s",
                          rmse_mm3, rmse_holt, meilleure),
       x = "Date", y = "TMP (%)") +
  theme_minimal(base_size = 11) +
  theme(legend.position = "top",
        plot.title = element_text(face = "bold"))
  

return(list(
  label      = label,
  df_plot    = df_plot,
  graph      = g,
  rmse_mm3   = rmse_mm3,
  rmse_holt  = rmse_holt,
  mae_mm3    = mae_mm3,
  mae_holt   = mae_holt,
  lb_mm3_p   = lb_mm3$p.value,
  lb_holt_p  = lb_holt$p.value,
  holt_fit   = holt_fit,
  prev4      = prev4,
  alpha      = alpha_opt,
  beta       = beta_opt,
  meilleure  = meilleure
))
}
# ─── 1.3 Application aux segments clés ──────────────────────────────────────

# Segments prioritaires (séries les plus fournies pour l'analyse)
segments <- list(
  list(pays = "Cameroun", instr = "BTA", mat = "13s"),
  list(pays = "Cameroun", instr = "BTA", mat = "26s"),
  list(pays = "Cameroun", instr = "BTA", mat = "52s"),
  list(pays = "Cameroun", instr = "OTA", mat = "1a"),
  list(pays = "Cameroun", instr = "OTA", mat = "2a"),
  list(pays = "Cameroun", instr = "OTA", mat = "3a"),
  list(pays = "Cameroun", instr = "OTA", mat = "5a"),
  list(pays = "Cameroun", instr = "OTA", mat = "6a"),
  list(pays = "Cameroun", instr = "OTA", mat = "10a"),
  
  list(pays = "Congo",    instr = "BTA", mat = "13s"),
  list(pays = "Congo",    instr = "BTA", mat = "26s"),
  list(pays = "Congo",    instr = "OTA", mat = "2a"),
  list(pays = "Congo",    instr = "OTA", mat = "3a"),
  list(pays = "Congo",    instr = "OTA", mat = "4a"),
  list(pays = "Congo",    instr = "OTA", mat = "5a"),

  list(pays = "Gabon",    instr = "BTA", mat = "13s"),
  list(pays = "Gabon",    instr = "BTA", mat = "26s"),
  list(pays = "Gabon",    instr = "BTA", mat = "52s"),
  list(pays = "Gabon",    instr = "OTA", mat = "2a"),  
  list(pays = "Gabon",    instr = "OTA", mat = "3a"),
  list(pays = "Gabon",    instr = "OTA", mat = "4a"),
  list(pays = "Gabon",    instr = "OTA", mat = "5a"),
  list(pays = "Gabon",    instr = "OTA", mat = "7a")
)

resultats_lissage <- list()
graphiques_lissage <- list()

for (seg in segments) {
  cle   <- paste(seg$pays, seg$instr, seg$mat, sep = "_")
  serie <- construire_serie(seg$pays, seg$instr, seg$mat)
  label <- paste(seg$pays, seg$instr, seg$mat)
  res   <- lisser_serie(serie, label = label)
  if (!is.null(res)) {
    resultats_lissage[[cle]] <- res
    graphiques_lissage[[cle]] <- res$graph
  }
}


# ─── 1.4 Tableau de synthèse des méthodes de lissage ────────────────────────

synthese_lissage <- data.frame(
  Série      = names(resultats_lissage),
  RMSE_MM3   = sapply(resultats_lissage, function(x) round(x$rmse_mm3,  4)),
  RMSE_Holt  = sapply(resultats_lissage, function(x) round(x$rmse_holt, 4)),
  MAE_MM3    = sapply(resultats_lissage, function(x) round(x$mae_mm3,   4)),
  MAE_Holt   = sapply(resultats_lissage, function(x) round(x$mae_holt,  4)),
  LjungBox_MM3_p  = sapply(resultats_lissage, function(x) round(x$lb_mm3_p,  3)),
  LjungBox_Holt_p = sapply(resultats_lissage, function(x) round(x$lb_holt_p, 3)),
  Alpha_Holt = sapply(resultats_lissage, function(x) x$alpha),
  Beta_Holt  = sapply(resultats_lissage, function(x) x$beta),
  Meilleure  = sapply(resultats_lissage, function(x) x$meilleure),
  stringsAsFactors = FALSE
)

cat("\n\n  ══ Tableau de synthèse — Lissage ══\n")
print(synthese_lissage, row.names = FALSE)

# Affichage des 4 premiers graphiques
if (length(graphiques_lissage) >= 4) {
  gridExtra::grid.arrange(
    graphiques_lissage[[1]], graphiques_lissage[[2]],
    graphiques_lissage[[3]], graphiques_lissage[[4]],
    ncol = 2,
    top = "ÉTAPE 1 — Lissage des courbes de taux CEMAC"
  )
}


# ─── 1.5 Courbe des taux lissée — Cameroun (snapshot actuel) ────────────────

# Construction d'un snapshot de la courbe par maturité pour le Cameroun
snapshot_cmr <- tmp_long %>%
  filter(Pays == "Cameroun", Date >= as.Date("2025-01-01")) %>%
  group_by(Maturite) %>%
  summarise(TMP_recent = mean(TMP, na.rm = TRUE), .groups = "drop") %>%
  mutate(
    Annees = case_when(
      Maturite == "13s"  ~ 0.25,
      Maturite == "26s"  ~ 0.50,
      Maturite == "52s"  ~ 1.00,
      Maturite == "2a"   ~ 2.00,
      Maturite == "3a"   ~ 3.00,
      Maturite == "4a"   ~ 4.00,
      Maturite == "5a"   ~ 5.00,
      Maturite == "6a"   ~ 6.00,
      Maturite == "7a"   ~ 7.00,
      Maturite == "10a"  ~ 10.00,
      TRUE ~ NA_real_
    )
  ) %>%
  filter(!is.na(Annees)) %>%
  arrange(Annees)

cat("\n  Snapshot courbe des taux — Cameroun (2025-2026) :\n")
print(snapshot_cmr, digits = 4)

g_courbe_cmr <- ggplot(snapshot_cmr, aes(x = Annees, y = TMP_recent)) +
  geom_point(linewidth = 3, colour = "steelblue") + # Changed size to linewidth
  geom_line(colour = "steelblue", linewidth = 1) + # Changed size to linewidth
  geom_smooth(method = "loess", se = TRUE, colour = "tomato",
              fill = "tomato", alpha = 0.15) +
  labs(title    = "Courbe des taux — Cameroun (lissage Loess)",
       subtitle  = "Taux moyens pondérés 2025-2026 | Source : BEAC",
       x = "Maturité (années)", y = "TMP (%)") +
  scale_x_continuous(breaks = c(0.25, 0.5, 1, 2, 3, 5, 10)) +
  theme_minimal(base_size = 12)
print(g_courbe_cmr)

# Code R à insérer à la fin de votre ÉTAPE 1 pour sauvegarder les graphiques dans un PDF
pdf(file = "Rapport_Lissage_Courbes_Taux.pdf", width = 11, height = 8.5) # Format Paysage

# 1. Tracer le snapshot de la courbe globale du Cameroun
print(g_courbe_cmr)

# 2. Boucler sur tous les segments calculés pour exporter les graphiques de lissage
for (nom_cle in names(graphiques_lissage)) {
  cat("Ajout du graphique au PDF :", nom_cle, "\n")
  print(graphiques_lissage[[nom_cle]])
}

# Fermeture du périphérique et finalisation du fichier PDF
dev.off()
cat("✔ Le fichier 'Rapport_Lissage_Courbes_Taux.pdf' a été généré avec succès.\n")

# ─────────────────────────────────────────────────────────────────────────────
# EXPORT DES DONNÉES LISSÉES VERS EXCEL (À insérer à la fin de l'ÉTAPE 1)
# ─────────────────────────────────────────────────────────────────────────────

cat("\n>>> Préparation de l'export Excel des données de taux lissées...\n")

# 1. Reconstitution du tableau long de toutes les séries lissées
donnees_excel_longue <- data.frame()

for (cle in names(resultats_lissage)) {
  res_seg <- resultats_lissage[[cle]]
  seg_info <- strsplit(cle, "_")[[1]]
  
  # Construction du dataframe pour ce segment spécifique
  df_temp <- data.frame(
    Date        = res_seg$df_plot$Date,
    Pays        = seg_info[1],
    Instrument  = seg_info[2],
    Maturite    = seg_info[3],
    TMP_Brut    = res_seg$df_plot$Brut,
    MM3         = res_seg$df_plot$MM3,
    MM5         = res_seg$df_plot$MM5,
    Holt        = res_seg$df_plot$Holt,
    stringsAsFactors = FALSE
  )
  donnees_excel_longue <- rbind(donnees_excel_longue, df_temp)
}

# 2. Préparation de la liste des onglets pour writexl
list_onglets <- list(
  "Synthèse Métriques" = synthese_lissage,
  "Données Lissées"   = donnees_excel_longue
)

# 3. Écriture du fichier Excel
nom_fichier_excel <- "Donnees_Courbe_Taux_Lissees_R.xlsx"
writexl::write_xlsx(list_onglets, path = nom_fichier_excel)

cat(sprintf("✔ Fichier Excel exporté avec succès : '%s' (%d lignes de données).\n\n", 
            nom_fichier_excel, nrow(donnees_excel_longue)))

################################################################################
## ÉTAPE 2 — ÉCONOMÉTRIE UNIVARIÉE : MODÈLES GARCH
##
##  Séquence obligatoire :
##    1. Test de stationnarité ADF  → travailler sur r_t ou Δr_t
##    2. Test d'effets ARCH          → justifier le GARCH
##    3. Test de normalité JB        → choisir distribution Student-t
##    4. Estimation GARCH(1,1)       → volatilité conditionnelle
##    5. ARMA-GARCH sur flux passif  → modélisation des rachats
################################################################################

cat("\n═══════════════════════════════════════════════════════\n")
cat("  ÉTAPE 2 — MODÈLES GARCH\n")
cat("═══════════════════════════════════════════════════════\n\n")


# ─── 2.1 Fonction complète : pipeline de tests + GARCH ───────────────────────

pipeline_garch <- function(serie_df, label = "", dist_model = "std") {
  # serie_df : data.frame avec colonnes Date et TMP
  
  y <- serie_df$TMP
  n <- length(y)
  
  if (n < 8) {
    cat(sprintf("  [SKIP] %s : trop peu d'observations (%d).\n", label, n))
    return(NULL)
  }
  
  cat(sprintf("\n  ═══ %s (n = %d) ═══\n", label, n))
  
  # ── Test 1 : Stationnarité ADF ──────────────────────────────────────────
  adf_niv  <- tryCatch(adf.test(y),    error = function(e) list(p.value = NA))
  adf_diff <- tryCatch(adf.test(diff(y)), error = function(e) list(p.value = NA))
  
  cat(sprintf("  ADF niveaux     : p = %.4f  %s\n",
              adf_niv$p.value,
              ifelse(is.na(adf_niv$p.value), "",
                     ifelse(adf_niv$p.value < 0.05, "→ Stationnaire ✔",
                            "→ Racine unitaire — travail sur diff !"))))
  cat(sprintf("  ADF différences : p = %.4f  %s\n",
              adf_diff$p.value,
              ifelse(is.na(adf_diff$p.value), "",
                     ifelse(adf_diff$p.value < 0.05, "→ Stationnaire ✔",
                            "→ Non stationnaire même en diff"))))
  
  # Choisir la série de travail
  if (!is.na(adf_niv$p.value) && adf_niv$p.value < 0.05) {
    serie_travail <- y
    type_serie    <- "Niveaux"
  } else {
    serie_travail <- diff(y)
    type_serie    <- "Différences premières"
  }
  cat(sprintf("  → Série utilisée : %s\n", type_serie))
  
  # ── Test 2 : Modèle AR(1) pour extraire les résidus ────────────────────
  ar1 <- tryCatch(arima(serie_travail, order = c(1, 0, 0)),
                  error = function(e) NULL)
  if (is.null(ar1)) {
    cat("  [ATTENTION] Ajustement AR(1) échoué.\n")
    return(NULL)
  }
  residus_ar <- residuals(ar1)
  
  # ── Test 3 : Effets ARCH ────────────────────────────────────────────────
  lag_arch <- min(5, floor(length(residus_ar) / 3))
  arch_test <- tryCatch(FinTS::ArchTest(residus_ar, lags = lag_arch),
                        error = function(e) list(p.value = NA))
  
  cat(sprintf("  Test ARCH (lag=%d) : p = %.4f  %s\n",
              lag_arch, arch_test$p.value,
              ifelse(is.na(arch_test$p.value), "",
                     ifelse(arch_test$p.value < 0.05,
                            "→ Effets ARCH confirmés ✔ GARCH justifié",
                            "→ Pas d'effets ARCH significatifs"))))
  
  # ── Test 4 : Normalité Jarque-Bera ──────────────────────────────────────
  jb_test <- tryCatch(jarque.bera.test(residus_ar),
                      error = function(e) list(p.value = NA))
  cat(sprintf("  Jarque-Bera     : p = %.4f  %s\n",
              jb_test$p.value,
              ifelse(is.na(jb_test$p.value), "",
                     ifelse(jb_test$p.value < 0.05,
                            "→ Non normalité → distribution Student-t recommandée",
                            "→ Normalité non rejetée"))))
  
  distribution <- if (!is.na(jb_test$p.value) && jb_test$p.value < 0.05) {
    "std"  # Student-t
  } else {
    "norm" # Gaussienne
  }
  cat(sprintf("  → Distribution GARCH : %s\n", distribution))
  
  # ── Estimation GARCH(1,1) ────────────────────────────────────────────────
  spec_garch <- ugarchspec(
    variance.model   = list(model = "sGARCH", garchOrder = c(1, 1)),
    mean.model       = list(armaOrder = c(1, 0), include.mean = TRUE),
    distribution.model = distribution
  )
  
  fit_garch <- tryCatch(
    ugarchfit(spec = spec_garch, data = serie_travail, solver = "hybrid"),
    error = function(e) NULL
  )
  
  if (is.null(fit_garch)) {
    cat("  [ATTENTION] Estimation GARCH échouée — tentative ARCH(1).\n")
    spec_arch1 <- ugarchspec(
      variance.model   = list(model = "sGARCH", garchOrder = c(1, 0)),
      mean.model       = list(armaOrder = c(0, 0), include.mean = TRUE),
      distribution.model = "norm"
    )
    fit_garch <- tryCatch(
      ugarchfit(spec = spec_arch1, data = serie_travail, solver = "hybrid"),
      error = function(e) NULL
    )
    if (is.null(fit_garch)) {
      cat("  [ERREUR] Aucun modèle GARCH ajustable pour cette série.\n")
      return(NULL)
    }
  }
  
  coefs   <- coef(fit_garch)
  sigma_t <- as.numeric(sigma(fit_garch))
  
  # Extraction des paramètres clés
  omega <- if ("omega" %in% names(coefs)) coefs["omega"] else NA
  alpha <- if ("alpha1" %in% names(coefs)) coefs["alpha1"] else NA
  beta  <- if ("beta1" %in% names(coefs)) coefs["beta1"] else NA
  
  persist <- if (!is.na(alpha) && !is.na(beta)) alpha + beta else NA
  demi_vie <- if (!is.na(beta) && beta > 0 && beta < 1) {
    round(log(0.5) / log(beta), 1)
  } else NA
  
  cat(sprintf("  GARCH(1,1) résultats :\n"))
  cat(sprintf("    omega  = %.6f\n", omega))
  cat(sprintf("    alpha  = %.4f  (réactivité aux chocs)\n", alpha))
  cat(sprintf("    beta   = %.4f  (persistance)\n", beta))
  cat(sprintf("    alpha+beta = %.4f  %s\n",
              persist,
              ifelse(is.na(persist), "",
                     ifelse(persist < 1, "✔ Stationnaire",
                            "⚠️  Non stationnaire !"))))
  if (!is.na(demi_vie))
    cat(sprintf("    Demi-vie volatilité = %.1f périodes\n", demi_vie))
  
  # Critères d'information
  ic <- infocriteria(fit_garch)
  cat(sprintf("    AIC = %.4f | BIC = %.4f\n", ic[1], ic[2]))
  
  # Prévision de volatilité à 4 périodes
  fcast_garch <- tryCatch(
    ugarchforecast(fit_garch, n.ahead = 4),
    error = function(e) NULL
  )
  if (!is.null(fcast_garch)) {
    vol_prev <- as.numeric(sigma(fcast_garch))
    cat(sprintf("    Volatilité prévue (4 pas) : %.4f %.4f %.4f %.4f\n",
                vol_prev[1], vol_prev[2], vol_prev[3], vol_prev[4]))
  }
  
  # Graphique : volatilité conditionnelle
  df_vol <- data.frame(
    Date     = serie_df$Date[if (type_serie == "Différences premières") -1 else 1:n],
    Sigma_t  = sigma_t
  )
  
  g_vol <- ggplot(df_vol, aes(x = Date, y = Sigma_t)) +
    geom_line(colour = "tomato", size = 0.9) +
    geom_hline(aes(yintercept = mean(Sigma_t, na.rm = TRUE)),
               linetype = "dashed", colour = "grey40") +
    labs(title    = paste("Volatilité conditionnelle —", label),
         subtitle  = sprintf("GARCH(1,1) | α+β = %.4f | Demi-vie = %s",
                             persist, demi_vie),
         x = "Date", y = "σ_t estimée") +
    theme_minimal(base_size = 11) +
    theme(plot.title = element_text(face = "bold"))
  print(g_vol)
  
  return(list(
    label      = label,
    fit        = fit_garch,
    sigma_t    = sigma_t,
    omega      = omega,
    alpha      = alpha,
    beta       = beta,
    persist    = persist,
    demi_vie   = demi_vie,
    aic        = ic[1],
    bic        = ic[2],
    adf_p_niv  = adf_niv$p.value,
    adf_p_diff = adf_diff$p.value,
    arch_p     = arch_test$p.value,
    jb_p       = jb_test$p.value,
    type_serie = type_serie,
    fcast_vol  = if (!is.null(fcast_garch)) vol_prev else NA
  ))
}


# ─── 2.2 Application aux segments principaux ────────────────────────────────

segments_garch <- list(
  list(pays = "Cameroun", instr = "BTA", mat = "13s"),
  list(pays = "Cameroun", instr = "BTA", mat = "26s"),
  list(pays = "Cameroun", instr = "BTA", mat = "52s"),
  list(pays = "Cameroun", instr = "OTA", mat = "1a"),
  list(pays = "Cameroun", instr = "OTA", mat = "2a"),
  list(pays = "Cameroun", instr = "OTA", mat = "3a"),
  list(pays = "Cameroun", instr = "OTA", mat = "5a"),
  list(pays = "Cameroun", instr = "OTA", mat = "6a"),
  list(pays = "Cameroun", instr = "OTA", mat = "10a"),
  
  list(pays = "Congo",    instr = "BTA", mat = "13s"),
  list(pays = "Congo",    instr = "BTA", mat = "26s"),
  list(pays = "Congo",    instr = "OTA", mat = "2a"),
  list(pays = "Congo",    instr = "OTA", mat = "3a"),
  list(pays = "Congo",    instr = "OTA", mat = "4a"),
  list(pays = "Congo",    instr = "OTA", mat = "5a"),
  
  list(pays = "Gabon",    instr = "BTA", mat = "13s"),
  list(pays = "Gabon",    instr = "BTA", mat = "26s"),
  list(pays = "Gabon",    instr = "BTA", mat = "52s"),
  list(pays = "Gabon",    instr = "OTA", mat = "2a"),  
  list(pays = "Gabon",    instr = "OTA", mat = "3a"),
  list(pays = "Gabon",    instr = "OTA", mat = "4a"),
  list(pays = "Gabon",    instr = "OTA", mat = "5a"),
  list(pays = "Gabon",    instr = "OTA", mat = "7a")
)

resultats_garch <- list()

for (seg in segments_garch) {
  cle   <- paste(seg$pays, seg$instr, seg$mat, sep = "_")
  serie <- construire_serie(seg$pays, seg$instr, seg$mat)
  label <- paste(seg$pays, seg$instr, seg$mat)
  res   <- pipeline_garch(serie, label = label)
  if (!is.null(res)) resultats_garch[[cle]] <- res
}


# ─── 2.3 ARMA-GARCH sur les flux de passif ──────────────────────────────────

cat("\n  ── Modélisation ARMA-GARCH sur les flux passif ──\n")
cat("     (Flux nets hebdomadaires = variations de parts)\n\n")

flux_nets_clean <- na.omit(passif_df$Flux_net)
n_flux          <- length(flux_nets_clean)
cat(sprintf("  Flux nets disponibles : %d observations\n", n_flux))
cat(sprintf("  Min : %.2f | Max : %.2f | Moy : %.2f | Sd : %.2f\n",
            min(flux_nets_clean), max(flux_nets_clean),
            mean(flux_nets_clean), sd(flux_nets_clean)))

# Avec 15 observations, on utilise un modèle parcimonieux
# Test de stationnarité
adf_flux <- tryCatch(adf.test(flux_nets_clean),
                     error = function(e) list(p.value = NA))
cat(sprintf("  ADF flux nets : p = %.4f  %s\n",
            adf_flux$p.value,
            ifelse(is.na(adf_flux$p.value), "",
                   ifelse(adf_flux$p.value < 0.05, "→ Stationnaire ✔",
                          "→ Racine unitaire possible"))))

# Sélection du meilleur modèle ARIMA
modele_arima_flux <- tryCatch(
  auto.arima(flux_nets_clean, seasonal = FALSE, stepwise = TRUE,
             approximation = FALSE),
  error = function(e) {
    arima(flux_nets_clean, order = c(1, 0, 0))
  }
)

cat(sprintf("\n  Meilleur ARIMA identifié : %s\n",
            paste0("ARIMA(", paste(modele_arima_flux$arma[1:3],
                                   collapse = ","), ")")))
cat(sprintf("  AIC = %.4f\n", AIC(modele_arima_flux)))

residus_flux <- residuals(modele_arima_flux)

# Tests sur les résidus
arch_flux <- tryCatch(FinTS::ArchTest(residus_flux, lags = 3),
                      error = function(e) list(p.value = NA))
jb_flux   <- tryCatch(jarque.bera.test(residus_flux),
                      error = function(e) list(p.value = NA))

cat(sprintf("  Test ARCH résidus  : p = %.4f\n", arch_flux$p.value))
cat(sprintf("  Jarque-Bera résidus: p = %.4f\n", jb_flux$p.value))

# Prévisions ARIMA des flux (4 semaines)
prev_flux_arima <- forecast(modele_arima_flux, h = 4)
cat("\n  Prévisions ARIMA — Flux nets (4 semaines) :\n")
cat(sprintf("    S+1 : %.2f parts  (IC 95%% : [%.2f ; %.2f])\n",
            prev_flux_arima$mean[1],
            prev_flux_arima$lower[1, 2],
            prev_flux_arima$upper[1, 2]))
cat(sprintf("    S+2 : %.2f parts  (IC 95%% : [%.2f ; %.2f])\n",
            prev_flux_arima$mean[2],
            prev_flux_arima$lower[2, 2],
            prev_flux_arima$upper[2, 2]))
cat(sprintf("    S+3 : %.2f parts\n", prev_flux_arima$mean[3]))
cat(sprintf("    S+4 : %.2f parts\n", prev_flux_arima$mean[4]))

g_flux <- autoplot(prev_flux_arima) +
  labs(title    = "Prévisions ARIMA — Flux nets de parts",
       subtitle  = "Fonds en phase de collecte — Horizon 4 semaines",
       x = "Semaine", y = "Flux net (parts)") +
  theme_minimal()
print(g_flux)


# ─── 2.4 Tableau de synthèse GARCH ──────────────────────────────────────────

if (length(resultats_garch) > 0) {
  synthese_garch <- data.frame(
    Série      = names(resultats_garch),
    ADF_Niv_p  = sapply(resultats_garch, function(x) round(x$adf_p_niv,  3)),
    ARCH_p     = sapply(resultats_garch, function(x) round(x$arch_p,     3)),
    JB_p       = sapply(resultats_garch, function(x) round(x$jb_p,       3)),
    Alpha      = sapply(resultats_garch, function(x) round(x$alpha,      4)),
    Beta       = sapply(resultats_garch, function(x) round(x$beta,       4)),
    Persist    = sapply(resultats_garch, function(x) round(x$persist,    4)),
    DemiVie    = sapply(resultats_garch, function(x) x$demi_vie),
    AIC        = sapply(resultats_garch, function(x) round(x$aic,        3)),
    stringsAsFactors = FALSE
  )
  cat("\n\n  ══ Tableau de synthèse — GARCH ══\n")
  print(synthese_garch, row.names = FALSE)
}


################################################################################
## ÉTAPE 3 — ÉCONOMÉTRIE MULTIVARIÉE : MODÈLES VAR
##
##  Phase II du projet :
##    A. VAR National Cameroun (BTA 26s vs OTA 5a)
##    B. VAR Régional 3 pays (OTA 5a CMR / GAB / CON)
##    C. IRF : choc Gabon → Cameroun et Congo
##    D. FEVD : décomposition de la variance
################################################################################

cat("\n═══════════════════════════════════════════════════════\n")
cat("  ÉTAPE 3 — MODÈLES VAR\n")
cat("═══════════════════════════════════════════════════════\n\n")


# ─── 3.1 Préparation des données VAR ────────────────────────────────────────

# Fonction : agrégation trimestrielle des données irrégulières
# (permet d'aligner des séries de fréquences différentes)
agregation_trimestrielle <- function(pays, instrument, maturite) {
  s <- construire_serie(pays, instrument, maturite)
  if (nrow(s) < 4) return(NULL)
  s$Trimestre <- paste0(format(s$Date, "%Y-T"),
                        ceiling(as.integer(format(s$Date, "%m")) / 3))
  s %>%
    group_by(Trimestre) %>%
    summarise(TMP_moy = mean(TMP, na.rm = TRUE), .groups = "drop") %>%
    arrange(Trimestre)
}

# Séries trimestrisées
cmr_bta26_trim  <- agregation_trimestrielle("Cameroun", "BTA", "26s")
cmr_ota5a_trim  <- agregation_trimestrielle("Cameroun", "OTA", "5a")
gab_ota5a_trim  <- agregation_trimestrielle("Gabon",    "OTA", "5a")
con_ota5a_trim  <- agregation_trimestrielle("Congo",    "OTA", "5a")

cat(sprintf("  CMR BTA 26s (trim.) : %d obs\n", nrow(cmr_bta26_trim)))
cat(sprintf("  CMR OTA  5a (trim.) : %d obs\n", nrow(cmr_ota5a_trim)))
cat(sprintf("  GAB OTA  5a (trim.) : %d obs\n", nrow(gab_ota5a_trim)))
cat(sprintf("  CON OTA  5a (trim.) : %d obs\n", nrow(con_ota5a_trim)))


# ─── 3.4 ANALYSES BIVARIÉES INTER-PAYS (CMR, GAB, CON — OTA 5 ans) ──────────

# Configuration des paires de pays à analyser
paires_analyse <- list(
  list(p1 = "Cameroun", name1 = "CMR_OTA5a", df1 = cmr_ota5a_trim,
       p2 = "Gabon",    name2 = "GAB_OTA5a", df2 = gab_ota5a_trim),
  
  list(p1 = "Cameroun", name1 = "CMR_OTA5a", df1 = cmr_ota5a_trim,
       p2 = "Congo",    name2 = "CON_OTA5a", df2 = con_ota5a_trim),
  
  list(p1 = "Gabon",    name1 = "GAB_OTA5a", df1 = gab_ota5a_trim,
       p2 = "Congo",    name2 = "CON_OTA5a", df2 = con_ota5a_trim)
)

# Boucle d'estimation automatique sur les 3 paires
for (paire in paires_analyse) {
  cat(sprintf("  ═════════════════════════════════════════════════\n"))
  cat(sprintf("  ANALYSE BIVARIÉE : %s ↔ %s (OTA 5 ans)\n", paire$p1, paire$p2))
  cat(sprintf("  ═════════════════════════════════════════════════\n\n"))
  
  # 1. Alignement et fusion des données de la paire
  merged_biv <- merge(
    setNames(paire$df1, c("Trimestre", paire$name1)),
    setNames(paire$df2, c("Trimestre", paire$name2)),
    by = "Trimestre"
  )
  
  n_obs <- nrow(merged_biv)
  cat(sprintf("  → Nombre de trimestres communs exploitables : %d\n", n_obs))
  
  # Condition de robustesse : au moins 8 observations pour un VAR bivarié d'ordre 1
  if (n_obs >= 8) {
    Y_biv <- as.matrix(merged_biv[, c(paire$name1, paire$name2)])
    
    # 2. Sélection de l'ordre optimal du VAR (limité à p=1 ou p=2 pour préserver les degrés de liberté)
    sel_biv <- tryCatch(
      VARselect(Y_biv, lag.max = min(2, floor(n_obs / 4)), type = "const"),
      error = function(e) NULL
    )
    
    p_opt <- 1
    if (!is.null(sel_biv)) {
      p_opt <- max(1, as.numeric(sel_biv$selection["AIC(n)"]))
    }
    cat(sprintf("  → Ordre sélectionné (AIC) : p = %d\n", p_opt))
    
    # 3. Estimation du modèle VAR Bivarié
    var_biv <- tryCatch(
      VAR(Y_biv, p = p_opt, type = "const"),
      error = function(e) {
        cat("  [ATTENTION] Échec de l'estimation avec p optimal. Repli sur p=1.\n")
        tryCatch(VAR(Y_biv, p = 1, type = "const"), error = function(ex) NULL)
      }
    )
    
    if (!is.null(var_biv)) {
      # 4. Diagnostic de stabilité (Racines de la matrice compagnon)
      racines_biv <- roots(var_biv)
      stable_flag <- all(racines_biv < 1)
      cat(sprintf("  → Stabilité du système : %s (Racine max = %.4f)\n", 
                  ifelse(stable_flag, "✔ Stable", "⚠️ Instable !"), max(racines_biv)))
      
      # 5. Test de Causalité au sens de Granger (Sens de transmission du risque)
      cat("\n  [TEST DE CAUSALITÉ DE GRANGER]\n")
      
      # Test p1 cause p2
      causa_1to2 <- causality(var_biv, cause = paire$name1)
      cat(sprintf("    - H0: %s ne cause pas %s | p-value = %.4f %s\n",
                  paire$p1, paire$p2, causa_1to2$Granger$p.value,
                  ifelse(causa_1to2$Granger$p.value < 0.05, "★ (Contagion significative)", "")))
      
      # Test p2 cause p1
      causa_2to1 <- causality(var_biv, cause = paire$name2)
      cat(sprintf("    - H0: %s ne cause pas %s | p-value = %.4f %s\n",
                  paire$p2, paire$p1, causa_2to1$Granger$p.value,
                  ifelse(causa_2to1$Granger$p.value < 0.05, "★ (Contagion significative)", "")))
      
      # 6. Fonctions de Réponse Impulsionnelle (IRF) — Analyse de propagation
      cat("\n  [FONCTIONS DE RÉPONSE IMPULSIONNELLE]\n")
      
      # Choc pays 1 → pays 2
      irf_1to2 <- tryCatch(
        irf(var_biv, impulse = paire$name1, response = paire$name2, n.ahead = 8, boot = FALSE),
        error = function(e) NULL
      )
      # Choc pays 2 → pays 1
      irf_2to1 <- tryCatch(
        irf(var_biv, impulse = paire$name2, response = paire$name1, n.ahead = 8, boot = FALSE),
        error = function(e) NULL
      )
      
      if (!is.null(irf_1to2) && !is.null(irf_2to1)) {
        impact_1to2 <- sum(irf_1to2$irf[[1]], na.rm = TRUE)
        impact_2to1 <- sum(irf_2to1$irf[[1]], na.rm = TRUE)
        cat(sprintf("    - Impact cumulé à 8 trimestres d'un choc (%s → %s) : %+.4f pts de taux\n", 
                    paire$p1, paire$p2, impact_1to2))
        cat(sprintf("    - Impact cumulé à 8 trimestres d'un choc (%s → %s) : %+.4f pts de taux\n", 
                    paire$p2, paire$p1, impact_2to1))
      }
      
      # 7. Décomposition de la Variance de l'Erreur de Prévision (FEVD)
      fevd_biv <- tryCatch(fevd(var_biv, n.ahead = 8), error = function(e) NULL)
      if (!is.null(fevd_biv)) {
        cat("\n  [FEVD — Horizon T+8]\n")
        cat(sprintf("    - Part de variance de %s expliquée par %s : %.1f%%\n",
                    paire$name1, paire$name2, fevd_biv[[paire$name1]][8, paire$name2] * 100))
        cat(sprintf("    - Part de variance de %s expliquée par %s : %.1f%%\n",
                    paire$name2, paire$name1, fevd_biv[[paire$name2]][8, paire$name1] * 100))
      }
      
    } else {
      cat("  [ERREUR] Impossible d'estimer le modèle VAR pour cette paire.\n")
    }
    
  } else {
    cat(sprintf("  [SKIP] Historique commun insuffisant (%d trimestres) pour une estimation bivariée fiable.\n", n_obs))
  }
  cat("\n")
}



################################################################################
## ÉTAPE 4 — ÉVALUATION DU RISQUE ALM
##
##  4.1 Liquidity at Risk (LaR 99%)
##  4.2 Gap de liquidité (coupons vs rachats prévus)
##  4.3 Contribution par pays à la performance
################################################################################

cat("\n═══════════════════════════════════════════════════════\n")
cat("  ÉTAPE 4 — ÉVALUATION DU RISQUE ALM\n")
cat("═══════════════════════════════════════════════════════\n\n")


# ─── 4.1 Liquidity at Risk (LaR) ─────────────────────────────────────────────

cat("  ── 4.1 Liquidity at Risk (LaR 99%) ──\n\n")

#creation de la dataframe des rachats
type_operation=c("rachat","rachat","rachat","rachat","rachat","rachat","rachat","rachat","rachat","rachat")
nbr_parts_rachats=c(18.40,4.53, 7.32,4.57,4.51, 4.56, 27.68,5.42,4.54, 14.43)
date_rachat=c("05/02/2026","17/02/2026","26/02/2026", "05/03/2026","09/03/2026", "16/03/2026","09/04/2026" ,"29/04/2026","29/04/2026","07/05/2026")
df_rachats=data.frame(type_operation=type_operation,nbr_parts_rachats=nbr_parts_rachats,date_rachat=date_rachat)

moyenne_parts_rachats=mean(df_rachats$nbr_parts_rachats)
ecart_type_rachats=sd(df_rachats$nbr_parts_rachats)
LaR_Parts <- moyenne_parts_rachats + (ecart_type_rachats * 2.33)
LaR_Montant_FCFA   <- LaR_Parts * VL_PAR_PART_APPROX
cat(paste0(rep("-", 55), collapse = ""), "\n")
cat(sprintf("  🔥 LaR_Parts (Moyenne + 2.33*σ) : %.4f parts\n", LaR_Parts))
cat(sprintf("  💰 BESOIN DE LIQUIDITÉ EN FCFA  : %s FCFA\n", 
            format(round(LaR_Montant_FCFA), big.mark = " ")))
cat(paste0(rep("-", 55), collapse = ""), "\n")


# ─── 4.2 Gap de Liquidité ─────────────────────────────────────────────────────

# ─────────────────────────────────────────────────────────────────────────────
# ÉTAPE 4.2 — GAP DE LIQUIDITÉ (COUPONS ACQUIS VS SORTIES ESTIMÉES)
# ─────────────────────────────────────────────────────────────────────────────

cat("\n  ── 4.2 Gap de Liquidité (Croisement Actif/Passif) ──\n\n")

# =============================================================================
# ÉTAPE 1 : ÉVALUATION DES FLUX ENTRANTS DE L'ACTIF (COUPONS)
# =============================================================================
# Extraction du flux hebdomadaire moyen généré par l'encours obligataire actuel
actif$Coupon_hebdo <- (actif$Nominal_brut * actif$Taux_facial) / 52
actif$Coupon_hebdo[is.na(actif$Coupon_hebdo)] <- 0

coupon_total_hebdo  <- sum(actif$Coupon_hebdo, na.rm = TRUE)
coupon_total_annuel <- sum(actif$Nominal_brut * actif$Taux_facial, na.rm = TRUE)

cat(sprintf("    [Actif] Préciput de coupons hebdomadaires constants : %s FCFA\n",
            format(round(coupon_total_hebdo), big.mark = " ")))
cat(sprintf("    [Actif] Total des revenus de coupons courus annuels : %s FCFA\n",
            format(round(coupon_total_annuel), big.mark = " ")))

# =============================================================================
# ÉTAPE 2 : PARAMÉTRAGE DES SORTIES DU PASSIF (A PARTIR DE DF_RACHATS)
# =============================================================================
# Utilisation des métriques descriptives calculées sur votre dataframe dédiée
moyenne_parts_rachats <- mean(df_rachats$nbr_parts_rachats, na.rm = TRUE)
ecart_type_rachats    <- sd(df_rachats$nbr_parts_rachats, na.rm = TRUE)

# Rappel de la LaR paramétrique calculée à l'étape 4.1
LaR_Parts <- moyenne_parts_rachats + (ecart_type_rachats * 2.33)

# Conversion des volumes de parts en flux monétaires (FCFA)
# (Rachats exprimés en valeurs négatives pour matérialiser les décaissements)
flux_rachat_median_FCFA  <- -moyenne_parts_rachats * VL_PAR_PART_APPROX
flux_rachat_stresse_FCFA <- -LaR_Parts * VL_PAR_PART_APPROX

# =============================================================================
# ÉTAPE 3 : CONSTRUCTION DE LA MATRICE DES GAPS SUR 4 SEMAINES
# =============================================================================
semaines_horizon <- 1:4

gap_df <- data.frame(
  Semaine       = paste0("S+", semaines_horizon),
  Coupons_FCFA  = rep(coupon_total_hebdo, 4),
  Rachats_P50   = rep(flux_rachat_median_FCFA, 4),
  Rachats_P99   = rep(flux_rachat_stresse_FCFA, 4)
)

# Calcul des indicateurs de solde (Gap Net = Entrées - Abs(Sorties))
gap_df <- gap_df %>%
  mutate(
    Gap_Net_P50  = Coupons_FCFA - abs(Rachats_P50),
    Gap_Net_P99  = Coupons_FCFA - abs(Rachats_P99),
    Statut_P50   = ifelse(Gap_Net_P50 >= 0, "Excédentaire", "Déficitaire"),
    Statut_P99   = ifelse(Gap_Net_P99 >= 0, "Excédentaire", "Déficitaire")
  )

cat("\n    [MÉTRIQUES] Tableau de croisement ALM Horizon 4 semaines (en FCFA) :\n")
print(gap_df %>%
        dplyr::mutate(
          Coupons_FCFA = round(Coupons_FCFA),
          Rachats_P50  = round(Rachats_P50),
          Rachats_P99  = round(Rachats_P99),
          Gap_Net_P50  = round(Gap_Net_P50),
          Gap_Net_P99  = round(Gap_Net_P99)
        ))
sprintf("\n  ceci signifie que dans 4 semaines, le FCP devra trouver %f FCFA(cas extrême: 1 pourcent) pour honnorer leurs engagements futurs ",gap_df$Gap_Net_P99[1]) 

# =============================================================================
# ÉTAPE 4 : MODÉLISATION GRAPHIQUE COMPARÉE DES GAPS DE TRÉSORERIE (CORRIGÉE)
# =============================================================================


g_gap <- ggplot(gap_df, aes(x = Semaine)) +
  # Graphique du fond : Entrées stables de l'Actif (Coupons)
  geom_col(aes(y = Coupons_FCFA, fill = "Coupons entrants (Actif)"), alpha = 0.8, width = 0.5) +
  
  # Graphique de l'avant-plan : Scénario normal (P50)
  geom_col(aes(y = Rachats_P50, fill = "Rachats attendus (P50)"), alpha = 0.6, width = 0.4) +
  
  # Graphique de l'avant-plan : Scénario de crise de rachat (LaR P99)
  # Correction : Utilisation d'un code Hexadécimal sécurisé "#DC143C" à la place de "crimson"
  geom_col(aes(y = Rachats_P99, fill = "Rachats max stressés (LaR P99)"), alpha = 0.3, width = 0.4) +
  
  # Ligne de démarcation de l'équilibre de trésorerie (Solde Zéro)
  geom_hline(yintercept = 0, color = "black", size = 0.8) +
  
  # Paramétrage visuel des couleurs de la légende
  scale_fill_manual(values = c(
    "Coupons entrants (Actif)"       = "darkblue",
    "Rachats attendus (P50)"         = "orange",
    "Rachats max stressés (LaR P99)" = "#DC143C" # Code Hex pour un rouge profond/crimson stable
  ), name = "Composantes ALM") +
  
  labs(
    title    = "Croisement ALM : Tombées de Coupons vs Risques de Rachat",
    subtitle = "Modélisation des Gaps cumulés sous scénario normal (P50) et extrême (P99)",
    x        = "Horizon temporel (Semaines)", 
    y        = "Flux financiers (FCFA)"
  ) +
  scale_y_continuous(labels = scales::comma) +
  theme_minimal() +
  theme(legend.position = "top")

print(g_gap)

# ─── 4.3 Performance et contribution par pays ───────────────────────────────

cat("\n  ── 4.3 Performance et contribution par pays ──\n\n")

# =============================================================================
# ÉTAPE 1 : REGROUPEMENT PAR ÉMETTEUR ET CALCUL DES POIDS DU PORTEFEUILLE
# =============================================================================
# On commence par nettoyer les espaces superflus dans le nom des émetteurs.
# Ensuite, on calcule la valorisation globale du portefeuille pour déterminer
# le poids exact de chaque pays dans l'allocation d'actifs.
perf_pays <- actif %>%
  mutate(Emetteur_grp = trimws(Emetteur)) %>%
  group_by(Emetteur_grp) %>%
  summarise(
    Valorisation_tot  = sum(Valorisation, na.rm = TRUE),
    Nominal_brut_tot  = sum(Nominal_brut, na.rm = TRUE),
    
    # ÉTAPE 2 : CALCUL DU TAUX FACIAL MOYEN PONDÉRÉ PAR PAYS
    # Le rendement d'un pays dépend de la taille de chaque ligne à sa valeur de marché (Valorisation)
    Taux_facial_moy   = weighted.mean(Taux_facial, Valorisation, na.rm = TRUE),
    Coupon_annuel_tot = sum(Nominal_brut * Taux_facial, na.rm = TRUE),
    N_titres          = n(),
    .groups = "drop"
  ) %>%
  
  # ÉTAPE 3 : CALCUL DE LA CONTRIBUTION INDIVIDUELLE À LA PERFORMANCE GLOBAL
  # Formule : Contribution = Poids du pays dans le portefeuille × Rendement moyen du pays
  mutate(
    Poids        = Valorisation_tot / sum(Valorisation_tot),
    Contrib_taux = Poids * Taux_facial_moy * 100  # Exprimé en pourcentage (%)
  ) %>%
  arrange(desc(Valorisation_tot))

# =============================================================================
# ÉTAPE 4 : INTÉGRATION DES DURATIONS MODIFIÉES (DM) ET SENSIBILITÉ TASSIQUE
# =============================================================================
# On associe à chaque pays sa Duration Modifiée indicative pour mesurer le risque de taux
durations_approx <- c(
  "TCHAD"                      = 2.5,
  "CONGO"                      = 1.8,
  "Républiq Centrafricaine"    = 3.5,
  "République Centrafricaine"  = 3.5,
  "GABON"                      = 3.0,
  "CAMEROUN"                   = 5.5
)

perf_pays$DM_approx <- sapply(perf_pays$Emetteur_grp, function(em) {
  match_em <- names(durations_approx)[sapply(names(durations_approx),
                                             function(k) grepl(toupper(k), toupper(em)))]
  if (length(match_em) > 0) durations_approx[match_em[1]] else 3.0
})

# Calcul de la sensibilité absolue par pays pour un choc de taux de 1 point de base (bp)
perf_pays$Sensibilite_bp <- perf_pays$DM_approx * perf_pays$Valorisation_tot / 10000

# =============================================================================
# ÉTAPE 5 : AFFICHAGE DU TABLEAU DE BORD DE PERFORMANCE
# =============================================================================
cat("    [TABLEAU] Répartition, Rendement et Contributions par zone émettrice :\n\n")
print(perf_pays %>%
        dplyr::select(Emetteur_grp, Valorisation_tot, Poids, Taux_facial_moy,
                      Contrib_taux, DM_approx) %>%
        dplyr::mutate(
          Valorisation_tot  = round(Valorisation_tot),
          Poids             = round(Poids * 100, 2),            # Affichage en %
          Taux_facial_moy   = round(Taux_facial_moy * 100, 3),  # Affichage en %
          Contrib_taux      = round(Contrib_taux, 4),           # Affichage en %
          DM_approx         = round(DM_approx, 2)
        ),
      row.names = FALSE)

# =============================================================================
# ÉTAPE 6 : AGRÉGATION DES MÉTRIQUES COMPOSITES DU FONDS
# =============================================================================
# Le taux composite est la somme exacte des contributions de chaque pays.
taux_composite <- sum(perf_pays$Contrib_taux, na.rm = TRUE)

# La duration modifiée globale est la moyenne des DM pondérée par les valorisations
dm_composite   <- weighted.mean(perf_pays$DM_approx, perf_pays$Valorisation_tot, na.rm = TRUE)

# Détermination du montant total sous gestion (encours obligataire global)
total_obligations_portefeuille <- sum(perf_pays$Valorisation_tot, na.rm = TRUE)

cat(sprintf("  📊 RENDEMENT COMPOSITE DU PORTEFEUILLE : %.4f%%\n", taux_composite))
cat(sprintf("  ⏳ DURATION MODIFIÉE COMPOSITE        : %.2f ans\n", dm_composite))
cat(sprintf("  📉 SENSIBILITÉ DU FONDS À 1 BP (FCFA)   : %s FCFA\n",
            format(round(dm_composite * total_obligations_portefeuille / 10000), big.mark = " ")))
cat(paste0(rep("-", 65), collapse = ""), "\n\n")

# =============================================================================
# ÉTAPE 7 : MODÉLISATION GRAPHIQUE DE LA CONTRIBUTION
# =============================================================================
library(ggplot2)

g_contrib <- ggplot(perf_pays, aes(x = reorder(Emetteur_grp, Contrib_taux), y = Contrib_taux)) +
  # Création des barres de performance colorées par pays
  geom_col(aes(fill = Emetteur_grp), show.legend = FALSE, width = 0.6) +
  
  # Ajout des étiquettes de valeurs exactes au bout de chaque barre
  geom_text(aes(label = sprintf(" %.3f%%", Contrib_taux)), hjust = -0.1, size = 3.5, fontface = "bold") +
  
  # Inversion des axes pour une lecture naturelle sous forme de classement (Benchmark)
  coord_flip() +
  
  # Thème visuel épuré
  theme_minimal() +
  scale_fill_brewer(palette = "Set2") + # Palette de couleurs professionnelles et distinctes
  labs(
    title    = "Structure de la Performance : Contribution par Pays Émetteur",
    subtitle = sprintf("Rendement embarqué global du portefeuille : %.4f%%", taux_composite),
    x        = "", 
    y        = "Fraction du rendement total géré (%)"
  ) +
  theme(
    plot.title   = element_text(face = "bold", size = 12),
    axis.text.y  = element_text(face = "bold", size = 10),
    panel.grid.minor = element_blank()
  )

print(g_contrib)

################################################################################
## ÉTAPE 5 — STRESS TESTING & CONFORMITÉ COSUMAF
##
##  5.1 Stress Test Double Choc
##      Choc Actif  : +150 bps sur la courbe des taux
##      Choc Passif : retrait de 20% des encours
##  5.2 Vérification des ratios COSUMAF (division des risques)
##  5.3 VaR Marché (paramétrique, 99%, 10 jours)
################################################################################

cat("\n═══════════════════════════════════════════════════════\n")
cat("  ÉTAPE 5 — STRESS TESTING & CONFORMITÉ COSUMAF\n")
cat("═══════════════════════════════════════════════════════\n\n")


# ─── 5.1 Choc Actif : Hausse de 150 bps ─────────────────────────────────────

cat("  ── 5.1 Choc Actif (+150 bps) ──\n\n")

DELTA_TAUX_BPS <- 150  # Choc de taux en points de base
DELTA_TAUX     <- DELTA_TAUX_BPS / 10000  # En décimal

# Duration modifiée par titre (approx : DM = maturité restante approx / (1+r))
# On utilise le taux facial comme proxy du taux de marché
actif$DM_approx <- actif$DM_approx <- NA
actif$DM_approx[grepl("02 ANS|2 ANS",  actif$Libelle, ignore.case = TRUE)] <- 1.85
actif$DM_approx[grepl("03 ANS|3 ANS",  actif$Libelle, ignore.case = TRUE)] <- 2.60
actif$DM_approx[grepl("04 ANS|4 ANS",  actif$Libelle, ignore.case = TRUE)] <- 3.50
actif$DM_approx[grepl("05 ANS|5 ANS",  actif$Libelle, ignore.case = TRUE)] <- 4.10
actif$DM_approx[grepl("06 ANS|6 ANS",  actif$Libelle, ignore.case = TRUE)] <- 4.80
actif$DM_approx[is.na(actif$DM_approx)] <- 2.50  # Défaut

# Calcul de la perte de valorisation par titre
actif$Perte_choc_actif <- -actif$DM_approx * DELTA_TAUX * actif$Valorisation

perte_totale_actif <- sum(actif$Perte_choc_actif, na.rm = TRUE)
perte_pct_actif    <- perte_totale_actif / total_oblig * 100

cat(sprintf("  Choc : +%d bps sur toute la courbe\n", DELTA_TAUX_BPS))
cat(sprintf("  Perte totale actif    : %s FCFA\n",
            format(round(perte_totale_actif), big.mark = " ")))
cat(sprintf("  En %% de l'actif oblig : %.2f%%\n", perte_pct_actif))

# Détail par titre
actif_stress <- actif %>%
  dplyr::select(Libelle, Emetteur, Taux_facial, Valorisation, DM_approx,
                Perte_choc_actif) %>%
  dplyr::mutate(
    Taux_pct    = round(Taux_facial * 100, 2),
    Valo_M      = round(Valorisation / 1e6, 3),
    Perte_FCFA  = round(Perte_choc_actif),
    Perte_pct   = round(Perte_choc_actif / Valorisation * 100, 3)
  ) %>%
  arrange(Perte_FCFA)

cat("\n  Perte par titre (du plus affecté au moins affecté) :\n")
print(actif_stress %>%
        dplyr::select(Libelle, Emetteur, Taux_pct, Valo_M, DM_approx,
                      Perte_FCFA, Perte_pct),
      row.names = FALSE)

# Titre le plus vulnérable
pire_titre <- actif_stress$Libelle[1]
perte_max  <- actif_stress$Perte_FCFA[1]
cat(sprintf("\n  Titre le plus vulnérable : %s\n", pire_titre))
cat(sprintf("  Perte estimée          : %s FCFA\n",
            format(perte_max, big.mark = " ")))


# ─── 5.2 Choc Passif : Retrait de 20% ───────────────────────────────────────

cat("\n  ── 5.2 Choc Passif (retrait 20%) ──\n\n")

TAUX_RETRAIT <- 0.20  # 20% des encours retirés

parts_actuelles         <- max(parts_cumul)
parts_remboursees       <- TAUX_RETRAIT * parts_actuelles
montant_remboursement   <- parts_remboursees * VL_PAR_PART_APPROX

cat(sprintf("  Parts actuelles         : %.0f\n", parts_actuelles))
cat(sprintf("  Parts à rembourser (20%%): %.0f\n", parts_remboursees))
cat(sprintf("  Montant de remboursement: %s FCFA\n",
            format(round(montant_remboursement), big.mark = " ")))
cat(sprintf("  Espèces disponibles     : %s FCFA\n",
            format(ESPECES_FCFA, big.mark = " ")))

couverture_especes <- ESPECES_FCFA / montant_remboursement
cat(sprintf("  Ratio de couverture     : %.2fx  %s\n",
            couverture_especes,
            ifelse(couverture_especes >= 1, "✔ Couvert", "⚠️  Déficit !")))


# ─── 5.3 Double Choc combiné ─────────────────────────────────────────────────

cat("\n  ── 5.3 Scénario Double Choc (Actif + Passif) ──\n\n")

actif_initial  <- total_oblig + ESPECES_FCFA
actif_apres    <- actif_initial + perte_totale_actif - montant_remboursement
parts_restantes <- parts_actuelles - parts_remboursees
VL_apres_choc  <- actif_apres / parts_restantes
VL_avant_choc  <- actif_initial / parts_actuelles

variation_VL   <- (VL_apres_choc - VL_avant_choc) / VL_avant_choc * 100

cat(sprintf("  Actif initial           : %s FCFA\n",
            format(round(actif_initial), big.mark = " ")))
cat(sprintf("  Perte choc actif        : %s FCFA\n",
            format(round(perte_totale_actif), big.mark = " ")))
cat(sprintf("  Remboursement passif    : %s FCFA\n",
            format(round(-montant_remboursement), big.mark = " ")))
cat(sprintf("  Actif après double choc : %s FCFA\n",
            format(round(actif_apres), big.mark = " ")))
cat(sprintf("  VL avant choc           : %.2f FCFA/part\n", VL_avant_choc))
cat(sprintf("  VL après choc           : %.2f FCFA/part\n", VL_apres_choc))
cat(sprintf("  Variation de VL         : %+.2f%%\n", variation_VL))

# Déficit de trésorerie
deficit_tresorerie <- perte_totale_actif + montant_remboursement + ESPECES_FCFA
cat(sprintf("\n  Bilan trésorerie :\n"))
cat(sprintf("    Espèces dispo      : +%s FCFA\n",
            format(ESPECES_FCFA, big.mark = " ")))
cat(sprintf("    Pertes obligataires: %s FCFA\n",
            format(round(perte_totale_actif), big.mark = " ")))
cat(sprintf("    Remboursements     : -%s FCFA\n",
            format(round(montant_remboursement), big.mark = " ")))
solde_treso <- ESPECES_FCFA + perte_totale_actif - montant_remboursement
cat(sprintf("    Solde de trésorerie: %s FCFA  %s\n",
            format(round(solde_treso), big.mark = " "),
            ifelse(solde_treso >= 0, "✔ Positif", "⚠️  DÉFICIT !")))

# Quel titre céder en priorité pour faire face (moindre moins-value) ?
cat("\n  Stratégie de cession pour faire face aux rachats :\n")
actif$Poids_perte <- abs(actif$Perte_choc_actif) / actif$Valorisation
actif_cession <- actif %>%
  dplyr::mutate(Poids_pct = round(Poids_perte * 100, 2)) %>%
  arrange(Poids_perce = Poids_pct) %>%
  dplyr::select(Libelle, Emetteur, Valorisation, DM_approx, Poids_pct) %>%
  head(5)
cat("  Titres à céder en priorité (moindre sensibilité) :\n")
print(actif_cession, row.names = FALSE)


# ─── 5.4 VaR Marché Paramétrique ─────────────────────────────────────────────

cat("\n  ── 5.4 VaR de Marché (paramétrique, 99%, 10 jours) ──\n\n")

# VaR = DM × VaR_taux × Valeur_portefeuille
# On utilise la volatilité GARCH du segment CMR OTA 5a comme proxy
if ("Cameroun_OTA_5a" %in% names(resultats_garch)) {
  vol_garch_cmr <- tail(resultats_garch[["Cameroun_OTA_5a"]]$sigma_t, 1)
  cat(sprintf("  Volatilité GARCH CMR OTA5a (dernière estimée) : %.4f%%\n",
              vol_garch_cmr))
  
  # Annualisation (données trimestrielles → hebdomadaires approximatif)
  vol_10j         <- vol_garch_cmr / 100 * sqrt(10 / 52)
  z_99            <- qnorm(0.99)
  
  VaR_taux_absolu <- z_99 * vol_10j  # En décimal
  VaR_marche      <- dm_composite * VaR_taux_absolu * total_oblig
  VaR_pct         <- VaR_marche / actif_brut * 100
  
  cat(sprintf("  Volatilité 10 jours (annualisée) : %.4f\n", vol_10j))
  cat(sprintf("  z_{99%%}                           : %.4f\n", z_99))
  cat(sprintf("  VaR taux 10j 99%%                  : %.4f (%.2f bps)\n",
              VaR_taux_absolu, VaR_taux_absolu * 10000))
  cat(sprintf("  VaR Marché portefeuille           : %s FCFA\n",
              format(round(VaR_marche), big.mark = " ")))
  cat(sprintf("  VaR en %% de l'actif brut          : %.3f%%\n", VaR_pct))
} else {
  cat("  [INFO] Estimation GARCH CMR OTA5a non disponible — VaR historique.\n")
  VaR_marche <- 0.02 * total_oblig  # Proxy 2% valeur obligataire
  VaR_pct    <- VaR_marche / actif_brut * 100
  cat(sprintf("  VaR Marché (proxy 2%%) : %s FCFA\n",
              format(round(VaR_marche), big.mark = " ")))
}


# ─── 5.5 Conformité COSUMAF ──────────────────────────────────────────────────

cat("\n  ── 5.5 Conformité COSUMAF — Division des risques ──\n\n")

LIMITE_COSUMAF_EMETTEUR <- 0.35  # 35% max par émetteur
SEUIL_ALERTE            <- 0.28  # Orange dès 28%

conformite_df <- resume_emetteur %>%
  dplyr::mutate(
    Poids_decimal = Valorisation / sum(Valorisation),
    Poids_pct     = Poids_decimal * 100,
    Statut = case_when(
      Poids_decimal > LIMITE_COSUMAF_EMETTEUR ~ "🔴 ALERTE (> 35%)",
      Poids_decimal > SEUIL_ALERTE             ~ "🟡 SURVEILLANCE (> 28%)",
      TRUE                                      ~ "🟢 Conforme"
    ),
    Reduction_requise_FCFA = pmax(0,
                                  (Poids_decimal - LIMITE_COSUMAF_EMETTEUR) *
                                    sum(Valorisation))
  )

cat("  Tableau de conformité COSUMAF :\n")
print(conformite_df %>%
        dplyr::select(Emetteur_clean, Valorisation, Poids_pct, Statut,
                      Reduction_requise_FCFA) %>%
        dplyr::mutate(
          Valorisation           = format(round(Valorisation), big.mark = " "),
          Poids_pct              = round(Poids_pct, 2),
          Reduction_requise_FCFA = format(round(Reduction_requise_FCFA),
                                          big.mark = " ")
        ),
      row.names = FALSE)

# Calcul du montant total à réduire sur l'entité hors limite
reduction_totale <- sum(conformite_df$Reduction_requise_FCFA, na.rm = TRUE)
cat(sprintf("\n  Montant total à rééquilibrer : %s FCFA\n",
            format(round(reduction_totale), big.mark = " ")))

if (reduction_totale > 0) {
  cat("  Plan de rééquilibrage recommandé :\n")
  cat("  → Cession progressive de la surexposition RCA lors des prochaines\n")
  cat("    adjudications favorables. Réinvestissement sur Cameroun et Gabon.\n")
  cat("  → Objectif : poids RCA < 35%, poids Cameroun > 15%.\n")
}

# Graphique COSUMAF
g_cosumaf <- ggplot(conformite_df,
                    aes(x = reorder(Emetteur_clean, -Poids_pct),
                        y = Poids_pct, fill = Statut)) +
  geom_col(alpha = 0.85) +
  geom_hline(yintercept = 35, colour = "red",    linetype = "dashed",
             size = 1.2) +
  geom_hline(yintercept = 28, colour = "orange", linetype = "dashed",
             size = 0.8) +
  annotate("text", x = 0.5, y = 36, label = "Limite COSUMAF 35%",
           colour = "red", hjust = 0, size = 3.5) +
  scale_fill_manual(
    values = c("🔴 ALERTE (> 35%)"       = "#CC3333",
               "🟡 SURVEILLANCE (> 28%)" = "#FF9900",
               "🟢 Conforme"             = "#339933"),
    name = "Statut"
  ) +
  labs(title    = "Conformité COSUMAF — Division des risques par émetteur",
       subtitle  = "Règlement COSUMAF 23 mai 2023 — Limite 35% par émetteur",
       x = "Émetteur", y = "Poids (%)") +
  theme_minimal() +
  theme(legend.position = "bottom",
        plot.title = element_text(face = "bold"))
print(g_cosumaf)


################################################################################
## ÉTAPE 6 — EXPORT DES RÉSULTATS POUR L'OUTIL VBA
##
##  Production de fichiers CSV lisibles par le module ImporterCoefficientsR()
##  du COSUMAF Risk Monitor Excel.
################################################################################

cat("\n═══════════════════════════════════════════════════════\n")
cat("  ÉTAPE 6 — EXPORT DES RÉSULTATS (VBA Interface)\n")
cat("═══════════════════════════════════════════════════════\n\n")


# ─── 6.1 Export coefficients GARCH ───────────────────────────────────────────

if (length(resultats_garch) > 0) {
  export_garch <- data.frame(
    Parametre     = c("omega", "alpha1", "beta1", "persistance",
                      "demi_vie_periodes", "AIC", "BIC"),
    stringsAsFactors = FALSE
  )
  
  for (nm in names(resultats_garch)) {
    g <- resultats_garch[[nm]]
    export_garch[[nm]] <- c(
      g$omega, g$alpha, g$beta, g$persist, g$demi_vie, g$aic, g$bic
    )
  }
  
  write.csv(export_garch, "coefficients_garch.csv", row.names = FALSE)
  cat("  ✔ Fichier créé : coefficients_garch.csv\n")
  cat("    (À importer via le bouton 'ImporterCoefficientsR' dans Excel)\n")
}


# ─── 6.2 Export résultats ALM ────────────────────────────────────────────────

export_alm <- data.frame(
  Indicateur = c(
    "LaR_parts_P99",
    "LaR_montant_FCFA",
    "VL_par_part_FCFA",
    "Parts_actuelles",
    "Coupon_hebdo_total_FCFA",
    "Coupon_annuel_total_FCFA",
    "Taux_composite_portefeuille_pct",
    "Duration_Modifiee_composite_ans",
    "VaR_Marche_FCFA",
    "VaR_Marche_pct_actif"
  ),
  Valeur = c(
    round(LaR_parts),
    round(LaR_montant_FCFA),
    round(VL_PAR_PART_APPROX, 2),
    round(parts_actuelles),
    round(coupon_total_hebdo),
    round(coupon_total_annuel),
    round(taux_composite, 4),
    round(dm_composite, 4),
    round(VaR_marche),
    round(VaR_pct, 4)
  ),
  stringsAsFactors = FALSE
)

write.csv(export_alm, "coefficients_var.csv", row.names = FALSE)
cat("  ✔ Fichier créé : coefficients_var.csv\n")


# ─── 6.3 Export stress test ──────────────────────────────────────────────────

export_stress <- data.frame(
  Scenario         = c("Double_Choc_150bps_20pct"),
  Choc_actif_bps   = DELTA_TAUX_BPS,
  Choc_passif_pct  = TAUX_RETRAIT * 100,
  Perte_actif_FCFA = round(perte_totale_actif),
  Remboursement_FCFA = round(montant_remboursement),
  VL_avant         = round(VL_avant_choc, 2),
  VL_apres         = round(VL_apres_choc, 2),
  Variation_VL_pct = round(variation_VL, 4),
  Solde_treso_FCFA = round(solde_treso),
  stringsAsFactors = FALSE
)

write.csv(export_stress, "resultats_StressTest.csv", row.names = FALSE)
cat("  ✔ Fichier créé : resultats_StressTest.csv\n")


# ─── 6.4 Export poids COSUMAF (pour le Solveur VBA) ─────────────────────────

export_poids <- conformite_df %>%
  dplyr::select(Emetteur = Emetteur_clean,
                Valorisation_FCFA = Valorisation,
                Poids_actuel_pct  = Poids_pct,
                Statut,
                Reduction_requise_FCFA) %>%
  dplyr::mutate(
    Limite_COSUMAF_pct    = LIMITE_COSUMAF_EMETTEUR * 100,
    Taux_facial_moy_pct   = perf_pays$Taux_facial_moy[
      match(Emetteur, perf_pays$Emetteur_grp)] * 100,
    DM_approx             = perf_pays$DM_approx[
      match(Emetteur, perf_pays$Emetteur_grp)]
  )

write.csv(export_poids, "resultats_COSUMAF_Poids.csv", row.names = FALSE)
cat("  ✔ Fichier créé : resultats_COSUMAF_Poids.csv\n")


# ─── 6.5 Export prévisions de flux (pour le Dashboard VBA) ─────────────────

export_flux <- data.frame(
  Semaine     = paste0("S+", 1:4),
  Flux_P50    = round(as.numeric(prev_flux_arima$mean), 2),
  Flux_LO_95  = round(as.numeric(prev_flux_arima$lower[, 2]), 2),
  Flux_HI_95  = round(as.numeric(prev_flux_arima$upper[, 2]), 2),
  En_FCFA_P50 = round(as.numeric(prev_flux_arima$mean) * VL_PAR_PART_APPROX),
  stringsAsFactors = FALSE
)

write.csv(export_flux, "resultats_Flux_Previsions.csv", row.names = FALSE)
cat("  ✔ Fichier créé : resultats_Flux_Previsions.csv\n")
cat("\n  → Ces 5 fichiers CSV sont prêts à être importés dans\n")
cat("    le COSUMAF Risk Monitor via le bouton 'ImporterCoefficientsR'.\n")
