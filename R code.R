
# R version 4.4.2
# R studio version 2024.09.0-375

# library(tidyverse)
# library(MASS)
# library(dplyr)
# library(gtsummary)
# library(autoReg)
# library(rrtable)
# library(forestplot)
# library(forestploter)
# library(ggplot2)


# Read the data ####
data <- readxl::read_xlsx("data.xlsx")


# Convert variables to factors ####
data[, c(3,4, 6:21)] <- lapply(data[, c(3,4, 6:21)], as.factor)
data$group <- factor(data$group, levels = c("0", "1"), labels = c("soc", "pif"))
data$sex <- factor(data$sex, levels = c("1", "2"), labels = c("male", "female"))
data$area <- factor(data$area, levels = c("1","2"), labels = c("urban", "rural"))
data$maristatus <- factor(data$maristatus, levels = c("1", "2"), labels = c("married", "divorc or singl"))
data$education<- factor(data$education, levels = c("1", "2", "3", "4"), labels = c("college", "hschool", "mschool", "pschool"))
data$living <- factor(data$living, levels = c("1", "2"), labels = c("with spouse", "notwith spouse"))
data$distance <- factor(data$distance, levels = c("1", "2", "3"), labels = c("within 1 km", "1-3 km", "3 km or more"))
data$smoke <- factor(data$smoke, levels = c ("0", "1"), labels = c("no", "yes"))
data$drink <- factor(data$drink, levels = c ("0", "1"), labels = c("no", "yes"))
data$age_group <- factor(data$age_group, levels = c("1", "2", "3"), labels = c("60-65", "66-70", "≥71"))
data$occupation <- factor(data$occupation, levels = c("1", "2", "3"), labels = c("farmer", "retired", "others"))
data$income <- factor(data$income, levels = c("1", "2", "3", "4"), labels = c("[0,276)usd", "[276,414)usd", "[414,690)usd", "≥690usd"))
data$incomelevel <- factor(data$incomelevel, levels = c("1", "2"), labels = c("lower","higher"))
data$cdisease <- factor(data$cdisease, levels = c("0", "1"), labels = c("no", "yes"))
data$ref_success <- factor(data$ref_success, levels = c("0", "1"), labels = c("no", "yes"))


# Table 1 ####
gaze(~.,data = data,show.p=TRUE) |> table2docx()
gaze(group~.,data = data,show.p=TRUE) |> table2docx()


# Table 2 ####
fit <- glm(vaccination ~ group + age_group + sex + occupation + area + maristatus + education +income + living + distance, 
           data = data, family = binomial(link = "logit"))
tbl_merge(list(tbl_regression(fit, exponentiate = T)), 
          tab_spanner = c("aor all")) %>% as_flex_table() %>% flextable::save_as_docx(path = "aor.all.docx")
tbl_merge(list(tbl_regression(glm(vaccination ~ group, data = data, family = binomial(link = "logit")), exponentiate = T)), 
          tab_spanner = c("cor all")) %>% as_flex_table() %>% flextable::save_as_docx(path = "cor.all.docx")


# Figure 2 subgroup analysis ####

# incomelevel_aOR
fit.incomel <- glm(vaccination ~ group + age_group + sex + occupation + area + maristatus + 
                     education + living + distance, data = data[(data$incomelevel == "lower"), ], family = binomial(link = "logit"))
fit.incomeh <- glm(vaccination ~ group + age_group + sex + occupation + area + maristatus + 
                     education + living + distance, data = data[(data$incomelevel == "higher"), ], family = binomial(link = "logit"))
tbl_merge(list(
  tbl_regression(fit.incomel, exponentiate = T) |> modify_column_merge(pattern = "{estimate} ({conf.low}, {conf.high})", rows = !is.na(estimate)),
  tbl_regression(fit.incomeh, exponentiate = T) |> modify_column_merge(pattern = "{estimate} ({conf.low}, {conf.high})", rows = !is.na(estimate))), 
  tab_spanner = c("aor income = lower", "aor income = higher")) %>% as_flex_table() %>% flextable::save_as_docx(path = "aor.income.docx")
# incomelevel_cOR
tbl_merge(list(
  tbl_regression(glm(vaccination ~ group, data = data[data$incomelevel == "lower", ], family = binomial(link = "logit")), exponentiate = T) |>
    modify_column_merge(pattern = "{estimate} ({conf.low}, {conf.high})", rows = !is.na(estimate)),
  tbl_regression(glm(vaccination ~ group, data = data[data$incomelevel == "higher", ], family = binomial(link = "logit")), exponentiate = T) |>
    modify_column_merge(pattern = "{estimate} ({conf.low}, {conf.high})", rows = !is.na(estimate))), 
  tab_spanner = c("cor income = lower", "cor income = higher")) %>% as_flex_table() %>% flextable::save_as_docx(path = "cor.income.docx")

# living_aOR
fit.withspous<- glm(vaccination ~ group + age_group + sex + occupation + area + education + distance + income, 
                    data = data[(data$living == "with spouse"), ], family = binomial(link = "logit"))
fit.notwithspous <- glm(vaccination ~ group + age_group + sex + occupation + area +education + distance + income, 
                        data = data[(data$living == "notwith spouse"), ], family = binomial(link = "logit"))
tbl_merge(list(
  tbl_regression(fit.withspous, exponentiate = T) |> modify_column_merge(pattern = "{estimate} ({conf.low}, {conf.high})", rows = !is.na(estimate)),
  tbl_regression(fit.notwithspous, exponentiate = T) |> modify_column_merge(pattern = "{estimate} ({conf.low}, {conf.high})", rows = !is.na(estimate))), 
  tab_spanner = c("aor living = with spouse", "aor living = notwith spouse")) %>% as_flex_table() %>% flextable::save_as_docx(path = "aor.living.docx")
# living_cOR
tbl_merge(list(
  tbl_regression(glm(vaccination ~ group, data = data[(data$living == "with spouse"), ], family = binomial(link = "logit")), exponentiate = T) |>
    modify_column_merge(pattern = "{estimate} ({conf.low}, {conf.high})", rows = !is.na(estimate)),
  tbl_regression(glm(vaccination ~ group, data = data[(data$living == "notwith spouse"), ], family = binomial(link = "logit")), exponentiate = T) |>
    modify_column_merge(pattern = "{estimate} ({conf.low}, {conf.high})",rows = !is.na(estimate))), 
  tab_spanner = c("cor living = with spouse", "cor living = notwith spouse")) %>% as_flex_table() %>% flextable::save_as_docx(path = "cor.living.docx")

# chronic disease_aOR
fit.wcd <- glm(vaccination ~ group + age_group + sex + maristatus + occupation + area + education + distance + income + living,
               data = data[data$cdisease == "yes", ], family = binomial(link = "logit"))
fit.ncd <- glm(vaccination ~ group + age_group + sex + maristatus + occupation + area + education + distance + income + living, 
               data = data[data$cdisease == "no", ], family = binomial(link = "logit"))
tbl_merge(list(
  tbl_regression(fit.wcd, exponentiate = T),tbl_regression(fit.ncd, exponentiate = T)), 
  tab_spanner = c("aor cdisease = yes", "aor cdisease = no")) %>% as_flex_table() %>% flextable::save_as_docx(path = "aor.cdisease.docx")
# chronic disease_cOR
tbl_merge(list(
  tbl_regression(glm(vaccination ~ group, data = data[data$cdisease == "yes", ], family = binomial(link = "logit")), exponentiate = T),
  tbl_regression(glm(vaccination ~ group, data = data[data$cdisease == "no", ], family = binomial(link = "logit")), exponentiate = T)), 
  tab_spanner = c("cor cdisease = yes", "cor cdisease = no")) %>% as_flex_table() %>% flextable::save_as_docx(path = "cor.cdisease.docx")

# area_aOR
fit.urban <- glm(vaccination ~ group + age_group + sex + maristatus + occupation + education + distance + income + living, 
                 data = data[data$area == "urban", ], family = binomial(link = "logit"))
fit.rural <- glm(vaccination ~  group + age_group + sex + maristatus + occupation + education + distance + income + living,
                 data = data[data$area == "rural", ], family = binomial(link = "logit"))
tbl_merge(list(
  tbl_regression(fit.urban, exponentiate = T),tbl_regression(fit.rural, exponentiate = T)), 
  tab_spanner = c("aor area = urban", "aor area = rural")) %>% as_flex_table() %>% flextable::save_as_docx(path = "aor.area.docx")
# area_cOR
tbl_merge(list(
  tbl_regression(glm(vaccination ~ group, data = data[data$area == "urban", ], family = binomial(link = "logit")), exponentiate = T),
  tbl_regression(glm(vaccination ~ group, data = data[data$area == "rural", ], family = binomial(link = "logit")), exponentiate = T)), 
  tab_spanner = c("cor area = urban", "cor area = rural")) %>% as_flex_table() %>% flextable::save_as_docx(path = "cor.area.docx")


# Table 3 ####
table(data$group, data$ref_success)
chisq.test(data$group, data$ref_success) #P-value for successfully referred vaccine 


# Figure 3 confidence analysis ####

# Confidence in safety_aOR_cOR
fit.safety <- glm(c_safety ~ group + age_group + sex + maristatus + education + occupation + income + living + distance, 
                  data = data, family = binomial(link = "logit"))
tbl_merge(list(tbl_regression(fit.safety, exponentiate = T)), 
          tab_spanner = c("aor safety")) %>% as_flex_table() %>% flextable::save_as_docx(path = "aor.safety.docx")
tbl_merge(list(tbl_regression(glm(c_safety ~ group, data = data, family = binomial(link = "logit")), exponentiate = T)),
          tab_spanner = c("cor safety")) %>% as_flex_table() %>% flextable::save_as_docx(path = "cor.safety.docx")

# Confidence in manage_aOR_cOR
fit.manage <- glm(c_manage ~ group + age_group + sex + maristatus + education + occupation + income + living + distance, 
                  data = data, family = binomial(link = "logit"))
tbl_merge(list(tbl_regression(fit.manage, exponentiate = T)), 
          tab_spanner = c("aor manage")) %>% as_flex_table() %>% flextable::save_as_docx(path = "aor.manage.docx")
tbl_merge(list(tbl_regression(glm(c_manage ~ group, data = data, family = binomial(link = "logit")), exponentiate = T)
), tab_spanner = c("cor manage")) %>% as_flex_table() %>% flextable::save_as_docx(path = "cor.manage.docx")

# Confidence in importance_aOR_cOR
fit.importnc <- glm(c_importnc ~ group + age_group + sex + maristatus + education + occupation + income + living + distance, 
                    data = data, family = binomial(link = "logit"))
tbl_merge(list(tbl_regression(fit.importnc, exponentiate = T)), 
          tab_spanner = c("aor importnc")) %>% as_flex_table() %>% flextable::save_as_docx(path = "aor.importnc.docx")
tbl_merge(list(tbl_regression(glm(c_importnc ~ group, data = data, family = binomial(link = "logit")), exponentiate = T)), 
          tab_spanner = c("cor importnc")) %>% as_flex_table() %>% flextable::save_as_docx(path = "cor.importnc.docx")

# Confidence in effectiveness_aOR_cOR
fit.effect <- glm(c_effect ~ group + age_group + sex + maristatus + education + occupation + income + living + distance, 
                  data = data, family = binomial(link = "logit"))
tbl_merge(list(tbl_regression(fit.effect, exponentiate = T)), 
          tab_spanner = c("aor effect")) %>% as_flex_table() %>% flextable::save_as_docx(path = "aor.effect.docx")
tbl_merge(list(tbl_regression(glm(c_effect ~ group, data = data, family = binomial(link = "logit")), exponentiate = T)), 
          tab_spanner = c("cor effect")) %>% as_flex_table() %>% flextable::save_as_docx(path = "cor.effect.docx")










#  Plot - Figure 2 ####

# Read the forest data 
forestdata <- readxl::read_xlsx("data.figure2.xlsx")

# Prepare an empty forest plot
forestdata$`Crude Odds Ratio` <- paste0(rep(" ", 20), collapse = " ")
forestdata$`Adjusted Odds Ratio` <- paste0(rep(" ", 10), collapse = " ")
forestdata[, c(2:6)] <- lapply(forestdata[, c(2:6)], function(x) ifelse(is.na(x), " ", x))
colnames(forestdata)[c(4, 6)] <- c("P        ", "P")

# Define a custom forest plot theme
tm <- forest_theme(
  base_size = 10.5,
  ci_pch = 18,
  ci_col = "black",
  ci_lty = 1,
  ci_lwd = 1.8,
  ci_Theight = 0.3,
  refline_gp = gpar(lwd = 1, lty = "dashed", col = "grey20"),
  footnote_gp = gpar(cex = 1, fontface = "italic", col = "red4")
)

# Draw the forest plot
# P_cOR
forest(data = forestdata[, c(1:2, 13, 3:4)],
       est = list(forestdata$cOR), lower = list(forestdata$clower), upper = list(forestdata$cupper),
       ci_column = c(3), ref_line = 1, sizes = 0.5, xlim = c(0.5, 100), ticks_at = c(1, 20, 40, 60, 80),
       theme = tm) %>%
  edit_plot(part = "header", gp = gpar(fontsize = 10, fontface = "bold")) %>%
  add_border(part = "header", where = "bottom") %>%
  add_border(part = "header", where = "top") %>%
  edit_plot(part = "header", gp = gpar(fontface = "bold"),
            which = "text", col = which(colnames(forestdata[, c(1:2, 13, 3:4)]) == "P")) %>%
  edit_plot(part = "body", gp = gpar(fontface = "bold"),
            which = "text", row = which(forestdata[, 2] == " "), col = 1) %>%
  edit_plot(part = "body", which = "ci", row = which(forestdata$clower > 1),
            col = 3, gp = gpar(col = "#f48c37", lwd = 1.5)) -> P_cOR
# P_aOR
forest(data = forestdata[, c(1:2, 14, 5:6)],
       est = list(forestdata$cOR), lower = list(forestdata$clower), upper = list(forestdata$cupper),
       ci_column = c(3), ref_line = 1, sizes = 0.5, xlim = c(0.5, 100), ticks_at = c(1, 20, 40, 60, 80),
       theme = tm) %>%
  edit_plot(part = "header", gp = gpar(fontsize = 10, fontface = "bold")) %>%
  add_border(part = "header", where = "bottom") %>%
  add_border(part = "header", where = "top") %>%
  edit_plot(part = "header", gp = gpar(fontface = "bold"),
            which = "text", col = which(colnames(forestdata[, c(1:2, 14, 5:6)]) == "P")) %>%
  edit_plot(part = "body", gp = gpar(fontface = "bold"),
            which = "text", row = which(forestdata[, 2] == " "), col = 1) %>%
  edit_plot(part = "body", which = "ci", row = which(forestdata$clower > 1),
            col = 3, gp = gpar(col = "#f48c37", lwd = 1.5)) -> P_aOR
# library(gridExtra)
grid.arrange(P_cOR,P_aOR,ncol = 1, nrow = 2)



#  Plot - Figure 3 ####

# Read the forest data 
forsdat <- readxl::read_xlsx("data.figure3.xlsx")

# Prepare an empty forest plot
forsdat$`Crude Odds Ratio` <- paste0(rep(" ", 19), collapse = " ")
forsdat$`Adjusted Odds Ratio` <- paste0(rep(" ", 10), collapse = " ")
forsdat[, c(2:6)] <- lapply(forsdat[, c(2:6)], function(x) ifelse(is.na(x), " ", x))
colnames(forsdat)[c(4, 6)] <- c("P        ", "P")

# Define a custom forest plot theme
tm <- forest_theme(
  base_size = 10.5,
  ci_pch = 18,
  ci_col = "black",
  ci_lty = 1,
  ci_lwd = 1.8,
  ci_Theight = 0.3,
  refline_gp   = gpar(lwd = 1, lty = "dashed", col = "grey20"),  
  footnote_gp  = gpar(cex = 1, fontface = "italic", col = "red4")
)

# Draw the forest plot
# P_cOR
forest(data = forsdat[, c(1:2, 13, 3:4)],
       est = list(forsdat$cOR), lower = list(forsdat$clower), upper = list(forsdat$cupper),
       ci_column = c(3), ref_line = 1, sizes = 0.5, xlim = c(0, 25), ticks_at = c(1, 5, 10, 15, 20), theme = tm) %>%
  edit_plot(part = "header", gp = gpar(fontsize = 10, fontface = "bold")) %>%
  add_border(part = "header", where = "bottom") %>%
  add_border(part = "header", where = "top") %>%
  edit_plot(part = "header", gp = gpar(fontface = "bold"), which = "text", col = which(colnames(forsdat[, c(1:2, 13, 3:4)]) == "P")) %>%
  edit_plot(part = "body", gp = gpar(fontface = "bold"), which = "text", row = which(forsdat[, 2] == " "), col = 1) %>%
  edit_plot(col = 3, row = which(forsdat$clower > 1), gp = gpar(col = "#f48c37", lwd = 1.5), which = "ci", part = "body") -> P_cOR
# P_aOR
forest(data = forsdat[, c(1:2, 14, 5:6)],
       est = list(forsdat$cOR), lower = list(forsdat$clower), upper = list(forsdat$cupper),
       ci_column = c(3), ref_line = 1, sizes = 0.5, xlim = c(0, 25), ticks_at = c(1, 5, 10, 15, 20), theme = tm) %>%
  edit_plot(part = "header", gp = gpar(fontsize = 10, fontface = "bold")) %>%
  add_border(part = "header", where = "bottom") %>%
  add_border(part = "header", where = "top") %>%
  edit_plot(part = "header", gp = gpar(fontface = "bold"), which = "text", col = which(colnames(forsdat[, c(1:2, 14, 5:6)]) == "P")) %>%
  edit_plot(part = "body", gp = gpar(fontface = "bold"), which = "text", row = which(forsdat[, 2] == " "), col = 1) %>%
  edit_plot(col = 3, row = which(forsdat$clower > 1), gp = gpar(col = "#f48c37", lwd = 1.5), which = "ci", part = "body") -> P_aOR
# library(gridExtra)
grid.arrange(P_cOR,P_aOR,ncol = 1, nrow = 2)

