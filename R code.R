rm(list = ls(all.names = TRUE))

# Install necessary packages if not already installed
install.packages(c("readxl", "tidyverse", "cluster", "openxlsx"))

# Load the required libraries
library(readxl)      # Read Excel files
library(tidyverse)   # Data manipulation & visualization
library(cluster)     # Clustering methods
library(openxlsx)    # Export to Excel


# IMPORT DATA FROM EXCEL
df <- read_csv("C:/Users/varis/OneDrive/Desktop/Exeter file/Term 3/Data/2b. CSV raw dataset.csv")

# 1. DEFINE VARAIBLES
#-----------------------------------------------------------------------

############# DEFINE CLUSTER VARIABLES ##########################
clustering_vars <- c(
  # Perceived benefits of LLMs (attitudinal)
  "BenLLMWhich1", "BenLLMWhich2", "BenLLMWhich3", "BenLLMWhich4",
  # Concerns about LLMs (attitudinal)
  "ConLLMWhich01", "ConLLMWhich02", "ConLLMWhich03", "ConLLMWhich04",
  "ConLLMWhich05", "ConLLMWhich06", "ConLLMWhich07", "ConLLMWhich08",
  "BenLLM", "ConLLM"
)

############# DEFINE PROFILE VARIABLES ##########################

# Categorical variables (summarized as % by cluster)
cat_vars <- c(
  "Cur_GOR", "cur_urbanruraltot", "Cur_Ethnic6", "Cur_disact2", "Cur_IntUse3",
  "Cur_PartyID5", "Cur_IdentClass", "Cur_RelStat5", "Cur_Tenure5",
  "Cur_EconAct5", "Cur_RClassGp", "Cur_SubjInc", "Cur_HHType"
)

# Binary Digital Skills (0/1), can be averaged
digskills_vars <- paste0("DigSkills", sprintf("%02d", 1:14))

# AI Governance + other numeric profiling
numeric_vars <- c(
  "AIGov_MonAIRisk_q", "AIGov_DevAISafety_q", "AIGov_AccessAIInfo_q", "AIGov_StopAIProduct_q",
  digskills_vars
)

# LLM experience variables
exp_vars <- c("ExpLLM_SearchTool_q", "ExpLLM_EntFun_q", "ExpLLM_Educational_q",
              "ExpLLM_JobApplication_q", "ExpLLM_EverydayTasks_q", "ExpLLM_Guidance_q")

# Combine all into one profiling list
profiling_vars <- c(cat_vars, numeric_vars, exp_vars)

# Create a dataframe containing only the variables for clustering
df_features <- df %>% select(all_of(clustering_vars))


# Create a dataframe containing only the variables for clustering
df_features <- df %>% select(all_of(clustering_vars))

# 2. PRE-PROCESSING FOR CLUSTERING
#-----------------------------------------------------------------------

############# BenLLM RESCALE ##########################

# This code creates a new, clean variable called 'BenLLM_Score'.
# It flips the scale to be more intuitive and removes invalid answers.
df <- df %>%
  mutate(BenLLM_Score = case_when(
    BenLLM == 1 ~ 4,  # 'Very beneficial' becomes 4 (Most Positive)
    BenLLM == 2 ~ 3,  # 'Fairly beneficial' becomes 3
    BenLLM == 3 ~ 2,  # 'Not very beneficial' becomes 2
    BenLLM == 4 ~ 1,  # 'Not at all beneficial' becomes 1 (Least Positive)
    TRUE ~ NA_real_   # All other values (-11, -9, -8, -1, and 5) become NA
  ))

############# CONLLM  Rescale ##########################

df <- df %>%
  mutate(ConLLM_Score = case_when(
    ConLLM == 1 ~ 4,  # Very concerned ??? 4 (highest concern)
    ConLLM == 2 ~ 3,  # Somewhat concerned ??? 3
    ConLLM == 3 ~ 2,  # Not very concerned ??? 2
    ConLLM == 4 ~ 1,  # Not at all concerned ??? 1 (least concern)
    TRUE ~ NA_real_   # Don't know or missing values ??? NA
  ))

############# CONLLM Which Rescale ##########################

# Define list of concern variables
con_vars <- c("ConLLMWhich01", "ConLLMWhich02", "ConLLMWhich03", "ConLLMWhich04",
              "ConLLMWhich05", "ConLLMWhich06", "ConLLMWhich07", "ConLLMWhich08")

# Keep only valid binary responses: 0 or 1; convert others (e.g. 5, -1) to NA
df <- df %>%
  mutate(across(all_of(con_vars), ~ ifelse(. %in% c(0, 1), ., NA)))

############# CLEAN DEMOGRAPHIC ##########################

# --- Clean the Age Category Variable ---

df <- df %>%
  mutate(Cur_AgeCat = ifelse(Cur_AgeCat > 0, Cur_AgeCat, NA))

# --- Clean the Sex Variable ---

df <- df %>%
  mutate(Cur_Sex = ifelse(Cur_Sex > 0, Cur_Sex, NA))

# --- Clean the Income Variable ---
df <- df %>%
  mutate(Cur_HHIncome4_21 = ifelse(Cur_HHIncome4_21 > 0, Cur_HHIncome4_21, NA))

# --- Clean the Education Variable ---  
df <- df %>%
  mutate(ExpLLM_Educational_q = ifelse(ExpLLM_Educational_q > 0, ExpLLM_Educational_q, NA))



# Create a list of all the experience columns to change
exp_cols <- c("ExpLLM_SearchTool_q", "ExpLLM_EntFun_q", "ExpLLM_Educational_q",
              "ExpLLM_JobApplication_q", "ExpLLM_EverydayTasks_q", "ExpLLM_Guidance_q")


############# ExpLLM RESCALE ##########################

# Use mutate() and across() to apply the corrected reverse coding
df <- df %>%
  mutate(across(all_of(exp_cols), ~case_when(
    . == 2 ~ 5,  # 'Yes, use regularly' becomes 5 (Highest Experience)
    . == 1 ~ 4,  # 'Yes, used a few times' becomes 4 (High Experience)
    . == 3 ~ 3,  # 'I'm not sure' stays in the middle
    . == 4 ~ 2,  # 'No, but open to it' becomes 2 (Low Experience)
    . == 5 ~ 1,  # 'No, and wouldn't want to' becomes 1 (Lowest Experience)
    TRUE ~ NA_real_  # All other values (like -1, -8, etc.) become NA
  )))

# Show first few values of the rescaled experience variables
head(df %>% select(all_of(exp_cols)))

############# OTHER BEN, AIConcern AND AIDecideComf RESCALE ##########################

df <- df %>%
  mutate(across(c("BenFR", "BenCancer", "BenLoan", "BenWB", "BenChatbot", 
                  "BenRoboCare", "BenCar", "AIConcern", "AIDecideComf"),
                ~ case_when(
                  . == 1 ~ 4,   # Very beneficial ??? 4
                  . == 2 ~ 3,   # Fairly beneficial ??? 3
                  . == 3 ~ 2,   # Not very beneficial ??? 2
                  . == 4 ~ 1,   # Not at all beneficial ??? 1
                  . > 0 & . <= 4 ~ .,  # (optional redundancy)
                  TRUE ~ NA_real_     # All else (5, -1, -8, etc.) ??? NA
                )))

############# RESCALE AI HARM ##########################
aih_vars <- c("AIHarm_ContPromVio_q", "AIHarm_FalseInfo_q", "AIHarm_FinFrauds_q", "AIHarm_DeepfakeImages_q")

# Recode: Remove code 4 (unsure), reverse scale so 1 = never, 3 = many times
df <- df %>%
  mutate(across(all_of(aih_vars), ~ case_when(
    . == 1 ~ 3,  # Many times ??? 3
    . == 2 ~ 2,  # A few times ??? 2
    . == 3 ~ 1,  # Never ??? 1
    . == 4 ~ NA_real_,  # Unsure ??? NA
    TRUE ~ NA_real_
  )))

df <- df %>%
  mutate(AIHarmConcern_Score = case_when(
    AIHarmConcern == 1 ~ 4,  # Very concerned ??? 4 (highest concern)
    AIHarmConcern == 2 ~ 3,  # Fairly concerned ??? 3
    AIHarmConcern == 3 ~ 2,  # Not very concerned ??? 2
    AIHarmConcern == 4 ~ 1,  # Not at all concerned ??? 1 (lowest concern)
    TRUE ~ NA_real_          # Invalid codes (e.g., -1, -8, 5) ??? NA
  ))


############# DELETE 5 AI GOV ##########################
aigov_vars <- c("AIGov_MonAIRisk_q", "AIGov_DevAISafety_q", "AIGov_AccessAIInfo_q", "AIGov_StopAIProduct_q")

# Replace code 5 with NA (keep original 1-4 scale)
df <- df %>%
  mutate(across(all_of(aigov_vars), ~ ifelse(. == 5, NA, .)))

############# CLEANING ##########################

# Create clustering feature dataset
df_features <- df %>% select(all_of(clustering_vars))

# Remove rows with missing values in clustering variables
df_features <- na.omit(df_features)

# Replace all negative values in the dataset with NA
df <- df %>%
  mutate(across(everything(), ~ ifelse(. < 0, NA, .)))

# Also filter the original dataset to keep the same rows
df_filtered <- df %>% filter(row_number() %in% as.numeric(rownames(df_features)))

# 3. CLUSTERING
#-----------------------------------------------------------------------
clustering_vars <- c(
  # Perceived benefits of LLMs (attitudinal)
  "BenLLMWhich1", "BenLLMWhich2", "BenLLMWhich3", "BenLLMWhich4",
  # Concerns about LLMs (attitudinal)
  "ConLLMWhich01", "ConLLMWhich02", "ConLLMWhich03", "ConLLMWhich04",
  "ConLLMWhich05", "ConLLMWhich06", "ConLLMWhich07", "ConLLMWhich08",
  "BenLLM_Score", "ConLLM_Score"
)


# Calculate Gower's distance, which is ideal for mixed data types (numeric, categorical, binary).
# This is a key adaptation from the user-provided code which used Euclidean distance.
gower_dist <- daisy(df_features, metric = "gower")

# Perform hierarchical clustering using Ward's method
# This part of the logic is taken directly from the user-provided R code
hc_model <- hclust(gower_dist, method = "complete")

# Plot the dendrogram to visualize the clustering
plot(hc_model, cex = 0.6, hang = -1, main = "Cluster Dendrogram for LLM Users")

# To determine the number of clusters, an elbow plot can be useful
plot(sort(hc_model$height, decreasing = TRUE)[1:20], type = "b", main = "Elbow Plot", xlab = "Number of Clusters", ylab = "Height")

# Based on analysis, we cut the tree to form 5 distinct clusters
# This aligns with the proposal's goal of creating actionable segments
clusters <- cutree(hc_model, k = 4)

# Add cluster labels to data
df_final <- df_filtered %>%
  mutate(cluster = clusters)

# 4. SILOUTTE TEST
#-----------------------------------------------------------------------
# 4. Optional: silhouette check
library(cluster)
sil <- silhouette(cutree(hc_model, k = 4), gower_dist)
plot(sil, col = 1:4, border = NA, main = "Silhouette Plot (k=4)")
mean(sil[, 3])  # average silhouette width

# 5. SEGMENT PROFILING (REVISED TO IMITATE USER'S EXAMPLE)
#-----------------------------------------------------------------------


# Compute segment sizes
proportions <- table(df_final$cluster) / nrow(df_final)
percentages <- proportions * 100
print("Segment Sizes (%):")
print(percentages)

segment_summary <- df_final %>%
  group_by(cluster) %>%
  summarise(
    mean_BenLLM = mean(BenLLM_Score, na.rm = TRUE),
    mean_ConLLM = mean(ConLLM_Score, na.rm = TRUE),
    across(all_of(setdiff(clustering_vars, c("BenLLM","ConLLM"))),
           ~ mean(as.numeric(.), na.rm = TRUE))
  )
# Calculate means of demographic variables per cluster
demographic_summary <- df_final %>%
  group_by(cluster) %>%
  summarise(across(all_of(profiling_vars), ~ mean(as.numeric(.), na.rm = TRUE), .names = "mean_{col}"))

# Merge summaries
full_segment_summary <- left_join(segment_summary, demographic_summary, by = "cluster") %>%
  mutate(size_percent = as.vector(percentages)) %>%
  select(cluster, size_percent, everything())  # Reorder columns

# View full segment summary
print(full_segment_summary)



# 5. EXPORT SUMMARY TO EXCEL
#-----------------------------------------------------------------------
# Load the library for writing Excel files
library(openxlsx)

# Create a new, empty Excel workbook
wb <- createWorkbook()

# Add a worksheet to the workbook named "Segment Profile Summary"
addWorksheet(wb, "Segment Profile Summary")

# Write the 'full_segment_summary' dataframe to the new worksheet
writeData(wb, "Segment Profile Summary", full_segment_summary)

# !!! IMPORTANT: Change the path below to your desired output folder and file name.
output_file_path <- "C:/Users/varis/OneDrive/Desktop/Exeter file/Term 3/R generated file/Segment summary 4 Final.xlsx"

# Save the workbook to the specified file path
saveWorkbook(wb, file = output_file_path, overwrite = TRUE)

# Print a confirmation message to the console
print(paste("Successfully exported summary results to:", output_file_path))


library(openxlsx)
library(dplyr)


# 6. ADD PROFILE VARIABLE BACK TO CLUSTER
#-----------------------------------------------------------------------

########################### CLEAN DATA: Remove invalid (e.g. < 0) ##########################

df <- df %>%
  mutate(across(all_of(profiling_vars), ~ ifelse(. < 0, NA, .)))


########################### NUMERIC PROFILE SUMMARY ##########################

numeric_profile <- df_final %>%
  group_by(cluster) %>%
  summarise(across(all_of(numeric_vars), ~ mean(as.numeric(.), na.rm = TRUE), .names = "mean_{col}"))

# Add segment size
segment_sizes <- df_final %>%
  count(cluster) %>%
  mutate(size_percent = round(n / sum(n) * 100, 1)) %>%
  select(cluster, size_percent)

numeric_profile <- left_join(segment_sizes, numeric_profile, by = "cluster")


########################### CREATE EXCEL WORKBOOK & EXPORT ###########################

wb <- createWorkbook()

# Sheet 1: Numeric / binary summary
addWorksheet(wb, "Numeric Profile")
writeData(wb, "Numeric Profile", numeric_profile)

# Add each categorical variable as % table by cluster
for (var in cat_vars) {
  tbl <- table(df_final$cluster, df_final[[var]])
  prop_tbl <- round(prop.table(tbl, margin = 1) * 100, 1)
  
  sheet_name <- substr(var, 1, 31)
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, as.data.frame.matrix(prop_tbl), rowNames = TRUE)
}

# Add each EXP LLM variable as % table by cluster
for (var in exp_vars) {
  tbl <- table(df_final$cluster, df_final[[var]])
  prop_tbl <- round(prop.table(tbl, margin = 1) * 100, 1)
  
  sheet_name <- paste0("EXP_", substr(var, 1, 28))  # Avoid Excel sheet name limit
  addWorksheet(wb, sheet_name)
  writeData(wb, sheet_name, as.data.frame.matrix(prop_tbl), rowNames = TRUE)
}

# Save to Excel file
output_path <- "C:/Users/varis/OneDrive/Desktop/Exeter file/Term 3/R generated file/Add profile to Segment.xlsx"
saveWorkbook(wb, file = output_path, overwrite = TRUE)

cat(paste("??? Exported full segment profile with EXP LLM to:", output_path))



# 6. SIGNIFICANT TESTING

# ----------------------------
# 0) Packages
# ----------------------------
pkgs <- c("dplyr","tidyr","stringr","purrr","openxlsx","FSA")
new  <- pkgs[!(pkgs %in% rownames(installed.packages()))]
if (length(new)) install.packages(new, dependencies = TRUE)
invisible(lapply(pkgs, library, character.only = TRUE))

# ----------------------------
# 1) Variable sets (from your lists)
# ----------------------------

# Categorical profiling variables (nominal/ordinal with named categories)
cat_vars <- c(
  "Cur_GOR","cur_urbanruraltot","Cur_Ethnic6","Cur_disact2","Cur_IntUse3",
  "Cur_PartyID5","Cur_IdentClass","Cur_RelStat5","Cur_Tenure5",
  "Cur_EconAct5","Cur_RClassGp","Cur_SubjInc","Cur_HHType",
  # Demographic banded vars often treated as categorical
  "Cur_AgeCat","Cur_Sex","Cur_HHIncome4_21","Cur_HEdQual"
)

# Binary "which" dummies (0/1)
which_vars <- c(
  paste0("BenLLMWhich", 1:4),
  paste0("ConLLMWhich", sprintf("%02d",1:8))
)

# Binary digital skills (0/1)
digskills_vars <- paste0("DigSkills", sprintf("%02d", 1:14))

# LLM experience (ordinal codes; treat as categorical for distribution tests,
# optionally also test mean ranks)
exp_vars <- c(
  "ExpLLM_SearchTool_q","ExpLLM_EntFun_q","ExpLLM_Educational_q",
  "ExpLLM_JobApplication_q","ExpLLM_EverydayTasks_q","ExpLLM_Guidance_q"
)

# Likert attitude items (1-4; some with 5 = "Don't know")
likert_vars <- c(
  "BenFR","BenCancer","BenLoan","BenWB","BenChatbot","BenRoboCare","BenCar",
  "AIConcern","AIDecideComf"
)

# Governance (Likert-style)
gov_vars <- c("AIGov_MonAIRisk_q","AIGov_DevAISafety_q",
              "AIGov_AccessAIInfo_q","AIGov_StopAIProduct_q")

# Composites / continuous scores
score_vars <- c("BenLLM_Score","ConLLM_Score","AIHarmConcern_Score")

# AI harm frequency items (ordinal with "many/few/never/unsure")
harm_vars <- c("AIHarm_ContPromVio_q","AIHarm_FalseInfo_q",
               "AIHarm_FinFrauds_q","AIHarm_DeepfakeImages_q")

# Master lists by test type
numeric_vars_for_anova <- c(score_vars, likert_vars, gov_vars)         # treat as numeric
categorical_vars        <- unique(c(cat_vars, exp_vars, harm_vars))    # distribution tests
binary_vars             <- unique(c(which_vars, digskills_vars))       # 0/1 tests

# Keep only variables that actually exist
exists_in_df <- function(v) v[v %in% names(df_final)]
numeric_vars_for_anova <- exists_in_df(numeric_vars_for_anova)
categorical_vars        <- exists_in_df(categorical_vars)
binary_vars             <- exists_in_df(binary_vars)

# ----------------------------
# 2) Cleaning
# ----------------------------
# Codes < 0 => NA
all_interest <- unique(c(numeric_vars_for_anova, categorical_vars, binary_vars))
df_final <- df_final %>%
  mutate(across(all_of(all_interest), ~ ifelse(. < 0, NA, .)))

# Likert items where 5 == "Don't know" -> NA
dk_vars <- intersect(likert_vars, names(df_final))
df_final[dk_vars] <- lapply(df_final[dk_vars], function(x) replace(x, x == 5, NA))

# Ensure cluster is factor
df_final$cluster <- as.factor(df_final$cluster)

# For safety, coerce strict 0/1 for binary sets
df_final[binary_vars] <- lapply(df_final[binary_vars], function(x) {
  y <- ifelse(x %in% c(0,1), x, NA)
  as.integer(y)
})

# ----------------------------
# 3) Helpers: effect sizes & tidiers
# ----------------------------
eta2_from_aov <- function(aov_obj){
  a <- anova(aov_obj)
  if (nrow(a) < 2) return(NA_real_)
  ss_between <- a$`Sum Sq`[1]
  ss_within  <- a$`Sum Sq`[2]
  ss_total   <- ss_between + ss_within
  if (!is.finite(ss_total) || ss_total == 0) return(NA_real_)
  ss_between / ss_total
}

epsilon2_from_kruskal <- function(kw, n, k){
  # kw$statistic is H; ??² = (H - k + 1)/(n - k)
  H <- as.numeric(kw$statistic)
  (H - k + 1) / (n - k)
}

cramers_v <- function(chi){
  chi2 <- as.numeric(chi$statistic)
  n    <- sum(chi$observed)
  r    <- nrow(chi$observed); c <- ncol(chi$observed)
  denom <- n * (min(r - 1, c - 1))
  if (denom <= 0) return(NA_real_)
  sqrt(chi2 / denom)
}

tidy_tukey <- function(tuk, var){
  out <- lapply(names(tuk), function(fac){
    as.data.frame(tuk[[fac]]) %>%
      mutate(contrast = rownames(tuk[[fac]]),
             variable = var, .before = 1)
  })
  bind_rows(out) %>%
    tidyr::separate(contrast, into = c("cluster1","cluster2"),
                    sep = "-", remove = FALSE)
}

# Pairwise proportions per single category level
pairwise_props_for_level <- function(tab, level, adjust = "holm"){
  successes <- tab[, level]
  totals    <- rowSums(tab)
  pw <- suppressWarnings(pairwise.prop.test(successes, totals, p.adjust.method = adjust))
  # Melt matrix into tidy frame
  if (all(is.na(pw$p.value))) return(NULL)
  mat <- pw$p.value
  L   <- rownames(mat); R <- colnames(mat)
  res <- list()
  for(i in seq_along(L)){
    for(j in seq_along(R)){
      if(!is.na(mat[i,j])){
        res[[length(res)+1]] <- data.frame(cluster1 = L[i], cluster2 = R[j], p_adj = mat[i,j])
      }
    }
  }
  if (!length(res)) return(NULL)
  out <- bind_rows(res)
  # Add proportions and difference
  props <- round(100 * successes / totals, 3)
  out <- out %>%
    mutate(prop1 = as.numeric(props[cluster1]),
           prop2 = as.numeric(props[cluster2]),
           diff_pp = prop1 - prop2)
  out
}

# ----------------------------
# 4) Run tests
# ----------------------------
overview_anova   <- list()
pairs_tukey      <- list()
overview_kruskal <- list()
pairs_dunn       <- list()   # for completeness on ordinal Likert (nonparametric)
means_by_cluster <- list()

# Numeric / Likert: ANOVA (+ Tukey), plus Kruskal (+ Dunn)
for (v in numeric_vars_for_anova){
  if (!v %in% names(df_final)) next
  dat <- df_final %>% select(cluster, !!sym(v)) %>% filter(!is.na(.data[[v]]))
  if (nrow(dat) == 0) next
  
  # means by cluster
  means_by_cluster[[v]] <- dat %>%
    group_by(cluster) %>%
    summarise(n = n(), mean = mean(.data[[v]], na.rm=TRUE), sd = sd(.data[[v]], na.rm=TRUE), .groups = "drop") %>%
    mutate(variable = v, .before = 1)
  
  # ANOVA
  m <- aov(dat[[v]] ~ dat$cluster)
  asu <- summary(m)[[1]]
  p   <- asu$`Pr(>F)`[1]
  Fv  <- asu$`F value`[1]
  df1 <- asu$Df[1]; df2 <- asu$Df[2]
  eta <- eta2_from_aov(m)
  overview_anova[[v]] <- data.frame(
    variable = v, test = "ANOVA", F = Fv, df1 = df1, df2 = df2,
    p_value = p, eta2 = eta
  )
  if (!is.na(p) && p < 0.05){
    tk <- TukeyHSD(m)
    tk_tidy <- tidy_tukey(tk, v) %>%
      rename(diff = diff, lwr = lwr, upr = upr, p_adj = `p adj`) %>%
      mutate(direction = ifelse(diff > 0, "cluster1 > cluster2", "cluster1 < cluster2"))
    pairs_tukey[[v]] <- tk_tidy
  }
  
  # Kruskal-Wallis + Dunn (ordinal-safe)
  kw <- kruskal.test(dat[[v]] ~ dat$cluster)
  k  <- nlevels(dat$cluster)
  n  <- nrow(dat)
  eps2 <- epsilon2_from_kruskal(kw, n, k)
  overview_kruskal[[v]] <- data.frame(
    variable = v, test = "Kruskal", H = as.numeric(kw$statistic),
    df = as.numeric(kw$parameter), p_value = kw$p.value, epsilon2 = eps2
  )
  if (!is.na(kw$p.value) && kw$p.value < 0.05){
    # Dunn with Holm
    dt <- suppressWarnings(FSA::dunnTest(dat[[v]] ~ dat$cluster, method = "holm"))
    pd <- dt$res %>%
      transmute(variable = v,
                cluster1 = gsub(" - .*","", Comparison),
                cluster2 = gsub(".*- ","", Comparison),
                Z = Z, p_adj = `P.adj`) %>%
      arrange(p_adj)
    pairs_dunn[[v]] <- pd
  }
}

overview_anova   <- bind_rows(overview_anova)
overview_kruskal <- bind_rows(overview_kruskal)
pairs_tukey      <- bind_rows(pairs_tukey)
pairs_dunn       <- bind_rows(pairs_dunn)
means_by_cluster <- bind_rows(means_by_cluster)

# ----------------------------
# Categorical (including exp_vars & harm_vars): ??² + per-level pairwise
# ----------------------------
chi_overview <- list()
prop_pairs   <- list()
cat_percent  <- list()

run_chi_for_var <- function(v){
  dat <- df_final %>% select(cluster, !!sym(v)) %>% filter(!is.na(.data[[v]]))
  if (nrow(dat) == 0) return(NULL)
  tab <- table(dat$cluster, dat[[v]])
  if (nrow(tab) < 2 || ncol(tab) < 2) return(NULL)
  
  cs <- suppressWarnings(chisq.test(tab))
  # Fisher if 2x2 and expected <5
  if (nrow(tab) == 2 && ncol(tab) == 2 && any(cs$expected < 5)){
    fs <- fisher.test(tab)
    cv <- cramers_v(cs) # Cramér's V from ??² (approx); fine for reporting
    chi_overview[[v]] <<- data.frame(
      variable = v, test = "Chi-square (2x2, Fisher p)", statistic = as.numeric(cs$statistic),
      df = as.numeric(cs$parameter), p_value = fs$p.value, cramers_v = cv
    )
  } else {
    cv <- cramers_v(cs)
    chi_overview[[v]] <<- data.frame(
      variable = v, test = "Chi-square", statistic = as.numeric(cs$statistic),
      df = as.numeric(cs$parameter), p_value = cs$p.value, cramers_v = cv
    )
  }
  
  # Row % table for interpretation
  pct <- round(prop.table(tab, margin = 1) * 100, 2)
  cat_percent[[v]] <<- data.frame(
    variable = v,
    cluster  = rownames(pct)[row(pct)],
    level    = colnames(pct)[col(pct)],
    percent  = as.numeric(pct)
  )
  
  # Attribute-level pairwise on EACH category
  levs <- colnames(tab)
  pp_all <- lapply(levs, function(L){
    out <- pairwise_props_for_level(tab, L, adjust = "holm")
    if (is.null(out)) return(NULL)
    out$variable <- v
    out$level    <- L
    out
  })
  prop_pairs[[v]] <<- bind_rows(pp_all)
  TRUE
}

invisible(lapply(categorical_vars, run_chi_for_var))
chi_overview <- bind_rows(chi_overview)
prop_pairs   <- bind_rows(prop_pairs)
cat_percent  <- bind_rows(cat_percent)

# ----------------------------
# Binary 0/1: ??² + pairwise on "1" (% Yes)
# ----------------------------
bin_overview <- list()
bin_pairs    <- list()
bin_percent  <- list()

for (v in binary_vars){
  if (!v %in% names(df_final)) next
  dat <- df_final %>% select(cluster, !!sym(v)) %>% filter(.data[[v]] %in% c(0,1))
  if (nrow(dat) == 0) next
  tab <- table(dat$cluster, dat[[v]])
  if (!all(c("0","1") %in% colnames(tab))) next
  
  cs <- suppressWarnings(chisq.test(tab))
  cv <- cramers_v(cs)
  bin_overview[[v]] <- data.frame(
    variable = v, test = "Chi-square (binary)", statistic = as.numeric(cs$statistic),
    df = as.numeric(cs$parameter), p_value = cs$p.value, cramers_v = cv
  )
  
  # row % 1s
  pct1 <- round(100 * tab[, "1"] / rowSums(tab), 2)
  bin_percent[[v]] <- data.frame(variable = v,
                                 cluster  = names(pct1),
                                 percent_yes = as.numeric(pct1))
  
  # pairwise on the "1" column
  pw <- suppressWarnings(pairwise.prop.test(tab[, "1"], rowSums(tab), p.adjust.method = "holm"))
  mat <- pw$p.value
  if (!all(is.na(mat))){
    L <- rownames(mat); R <- colnames(mat)
    res <- list()
    for(i in seq_along(L)){
      for(j in seq_along(R)){
        if(!is.na(mat[i,j])){
          res[[length(res)+1]] <- data.frame(cluster1=L[i], cluster2=R[j], p_adj=mat[i,j])
        }
      }
    }
    if (length(res)){
      bpp <- bind_rows(res)
      props <- round(100 * tab[, "1"] / rowSums(tab), 3)
      bpp <- bpp %>%
        mutate(prop1 = as.numeric(props[cluster1]),
               prop2 = as.numeric(props[cluster2]),
               diff_pp = prop1 - prop2,
               variable = v, level = "1")
      bin_pairs[[v]] <- bpp
    }
  }
}

bin_overview <- bind_rows(bin_overview)
bin_pairs    <- bind_rows(bin_pairs)
bin_percent  <- bind_rows(bin_percent)

# ---- Add 95% significance markers (???) before writing Excel ----
flag_star <- function(p) ifelse(!is.na(p) & p < 0.05, "???", "")

if (exists("overview_anova")   && nrow(overview_anova))   overview_anova   <- overview_anova   %>% dplyr::mutate(sig_95 = flag_star(p_value))
if (exists("overview_kruskal") && nrow(overview_kruskal)) overview_kruskal <- overview_kruskal %>% dplyr::mutate(sig_95 = flag_star(p_value))
if (exists("chi_overview")     && nrow(chi_overview))     chi_overview     <- chi_overview     %>% dplyr::mutate(sig_95 = flag_star(p_value))
if (exists("bin_overview")     && nrow(bin_overview))     bin_overview     <- bin_overview     %>% dplyr::mutate(sig_95 = flag_star(p_value))

if (exists("pairs_tukey")  && nrow(pairs_tukey))  pairs_tukey  <- pairs_tukey  %>% dplyr::mutate(sig_95 = flag_star(p_adj))
if (exists("pairs_dunn")   && nrow(pairs_dunn))   pairs_dunn   <- pairs_dunn   %>% dplyr::mutate(sig_95 = flag_star(p_adj))
if (exists("prop_pairs")   && nrow(prop_pairs))   prop_pairs   <- prop_pairs   %>% dplyr::mutate(sig_95 = flag_star(p_adj))
if (exists("bin_pairs")    && nrow(bin_pairs))    bin_pairs    <- bin_pairs    %>% dplyr::mutate(sig_95 = flag_star(p_adj))


# ----------------------------
# 5) Build workbook
# ----------------------------
wb <- createWorkbook()

# Overview sheets
if (nrow(overview_anova))   { addWorksheet(wb, "ANOVA_overview");       writeData(wb, "ANOVA_overview", overview_anova) }
if (nrow(overview_kruskal)) { addWorksheet(wb, "Kruskal_overview");     writeData(wb, "Kruskal_overview", overview_kruskal) }
if (nrow(chi_overview))     { addWorksheet(wb, "ChiSquare_overview");   writeData(wb, "ChiSquare_overview", chi_overview) }
if (nrow(bin_overview))     { addWorksheet(wb, "Binary_overview");      writeData(wb, "Binary_overview", bin_overview) }

# Means by cluster (for numeric vars)
if (nrow(means_by_cluster)) { addWorksheet(wb, "Means_by_cluster");     writeData(wb, "Means_by_cluster", means_by_cluster) }

# Pairwise results
if (nrow(pairs_tukey))      { addWorksheet(wb, "Tukey_pairs");          writeData(wb, "Tukey_pairs", pairs_tukey) }
if (nrow(pairs_dunn))       { addWorksheet(wb, "Dunn_pairs");           writeData(wb, "Dunn_pairs", pairs_dunn) }
if (nrow(prop_pairs))       { addWorksheet(wb, "Prop_pairs_by_level");  writeData(wb, "Prop_pairs_by_level", prop_pairs) }
if (nrow(bin_pairs))        { addWorksheet(wb, "Binary_pairs_yes");     writeData(wb, "Binary_pairs_yes", bin_pairs) }

# % tables for interpretation
if (nrow(cat_percent))      { addWorksheet(wb, "Cat_row_percents");     writeData(wb, "Cat_row_percents", cat_percent) }
if (nrow(bin_percent))      { addWorksheet(wb, "Binary_yes_percent");   writeData(wb, "Binary_yes_percent", bin_percent) }

# ----------------------------
# 6) Save
# ----------------------------
# Change to your desired path:
out_path <- "C:/Users/varis/OneDrive/Desktop/Exeter file/Term 3/R generated file/Significant test result.xlsx"
saveWorkbook(wb, out_path, overwrite = TRUE)
cat(paste0("\n??? Exported significance workbook to: ", out_path, "\n"))

# ----------------------------
# 7) Reading the outputs (quick guide)
# ----------------------------
# - ANOVA_overview: p_value < .05 => cluster means differ; use Tukey_pairs to see
#   which clusters differ and the direction (diff & 'direction' column).
# - Kruskal_overview / Dunn_pairs: nonparametric mirror for Likert.
# - ChiSquare_overview: distribution differs; see Prop_pairs_by_level to find
#   per-category which cluster has higher/lower % for that level (diff_pp sign).
# - Binary_overview / Binary_pairs_yes: the same logic on %Yes.

library(dplyr)

# Calculate cluster-level mean and SD for each variable
cluster_stats <- df_final %>%
  group_by(cluster) %>%
  summarise(
    n = n(),
    mean_BenLLM = mean(BenLLM, na.rm = TRUE),
    sd_BenLLM = sd(BenLLM, na.rm = TRUE),
    mean_ConLLM = mean(ConLLM, na.rm = TRUE),
    sd_ConLLM = sd(ConLLM, na.rm = TRUE),
    mean_EXP = mean(EXP_LLM, na.rm = TRUE),
    sd_EXP = sd(EXP_LLM, na.rm = TRUE)
  )

# Calculate overall mean and SD (for comparison)
overall_stats <- df_final %>%
  summarise(
    n = n(),
    mean_BenLLM = mean(BenLLM, na.rm = TRUE),
    sd_BenLLM = sd(BenLLM, na.rm = TRUE),
    mean_ConLLM = mean(ConLLM, na.rm = TRUE),
    sd_ConLLM = sd(ConLLM, na.rm = TRUE),
    mean_EXP = mean(EXP_LLM, na.rm = TRUE),
    sd_EXP = sd(EXP_LLM, na.rm = TRUE)
  )

# Print results
print(cluster_stats)
print(overall_stats)


############################ Mean vs Total mean significant ############################

# ------------------------------------------------------------
# CLUSTER vs TOTAL MEAN (one-sample t-tests)
#   Variables: BenLLM_Score, ConLLM_Score, and EACH EXP item
#   NOTE: Assumes your EXP items are already rescaled to 1..5 (as you did earlier).
# ------------------------------------------------------------

# 0) Packages
req <- c("dplyr","tidyr","purrr","broom","openxlsx","rlang")
new <- req[!(req %in% rownames(installed.packages()))]
if (length(new)) install.packages(new, dependencies = TRUE)
invisible(lapply(req, library, character.only = TRUE))

# 1) Variables
df_final$cluster <- as.factor(df_final$cluster)

# EXP items (already rescaled earlier in your pipeline)
exp_items <- c("ExpLLM_SearchTool_q","ExpLLM_EntFun_q","ExpLLM_Educational_q",
               "ExpLLM_JobApplication_q","ExpLLM_EverydayTasks_q","ExpLLM_Guidance_q")

score_vars <- c("BenLLM_Score","ConLLM_Score")
vars_all   <- c(score_vars, exp_items)

# Sanity checks
missing <- setdiff(vars_all, names(df_final))
if (length(missing)) stop("Missing in df_final: ", paste(missing, collapse=", "))

# 2) Helpers
one_sample_safe <- function(x, mu){
  x <- x[is.finite(x)]
  if (length(x) < 2 || !is.finite(mu)) {
    return(tibble(estimate = NA_real_, statistic = NA_real_, p.value = NA_real_,
                  conf.low = NA_real_, conf.high = NA_real_))
  }
  tt <- t.test(x, mu = mu)
  tibble(
    estimate  = unname(tt$estimate),
    statistic = unname(tt$statistic),
    p.value   = tt$p.value,
    conf.low  = tt$conf.int[1],
    conf.high = tt$conf.int[2]
  )
}

# 3) Overall descriptives (TOTAL row) for all vars
overall_desc <- df_final %>%
  summarise(across(
    all_of(vars_all),
    list(n = ~sum(is.finite(.)),
         mean = ~mean(., na.rm = TRUE),
         sd = ~sd(., na.rm = TRUE)),
    .names = "{.col}.{.fn}"
  )) %>%
  pivot_longer(everything(),
               names_to = c("variable",".value"),
               names_sep = "\\.") %>%
  rename(overall_n = n, overall_mean = mean, overall_sd = sd)

# 4) One-sample tests: each cluster mean vs its variable's TOTAL mean
run_one_sample <- function(var_names){
  purrr::map_dfr(var_names, function(vname){
    ov <- overall_desc %>% filter(variable == vname)
    ov_mu <- ov$overall_mean[1]; ov_sd <- ov$overall_sd[1]; ov_n <- ov$overall_n[1]
    
    purrr::map_dfr(levels(df_final$cluster), function(cl){
      x <- df_final %>% filter(cluster == cl) %>% pull(!!sym(vname))
      n_cl  <- sum(is.finite(x))
      m_cl  <- mean(x, na.rm = TRUE)
      sd_cl <- sd(x,   na.rm = TRUE)
      
      os <- one_sample_safe(x, ov_mu)
      
      # one-sample Cohen's d using cluster SD
      d1 <- if (is.finite(m_cl) && is.finite(sd_cl) && sd_cl > 0) (m_cl - ov_mu) / sd_cl else NA_real_
      
      tibble(
        variable         = vname,
        cluster          = as.character(cl),
        n_cluster        = n_cl,
        mean_cluster     = m_cl,
        sd_cluster       = sd_cl,
        overall_n        = ov_n,
        overall_mean     = ov_mu,
        overall_sd       = ov_sd,
        diff_vs_overall  = m_cl - ov_mu,
        pct_diff_overall = 100 * (m_cl - ov_mu) / ov_mu,
        t_stat           = os$statistic,
        p_raw            = os$p.value,
        ci_low           = os$conf.low,
        ci_high          = os$conf.high,
        cohens_d_1s      = d1
      )
    }) %>%
      group_by(variable) %>%
      mutate(p_holm = p.adjust(p_raw, method = "holm"),
             sig95  = ifelse(!is.na(p_holm) & p_holm < 0.05, "???", "")) %>%
      ungroup()
  }) %>% arrange(variable, cluster)
}

results_scores <- run_one_sample(score_vars)     # BenLLM_Score & ConLLM_Score
results_exp    <- run_one_sample(exp_items)      # Each EXP item separately
results_all    <- bind_rows(results_scores, results_exp)

# 5) Cluster descriptives + TOTAL row (for all variables)
cluster_desc <- df_final %>%
  group_by(cluster) %>%
  summarise(across(
    all_of(vars_all),
    list(n = ~sum(is.finite(.)),
         mean = ~mean(., na.rm = TRUE),
         sd = ~sd(., na.rm = TRUE)),
    .names = "{.col}_{.fn}"
  ), .groups = "drop")

total_row <- tibble(
  cluster = "TOTAL"
) %>%
  bind_cols(
    tibble(!!!setNames(lapply(vars_all, function(v) sum(is.finite(df_final[[v]]))), paste0(vars_all,"_n"))),
    tibble(!!!setNames(lapply(vars_all, function(v) mean(df_final[[v]], na.rm = TRUE)), paste0(vars_all,"_mean"))),
    tibble(!!!setNames(lapply(vars_all, function(v) sd(df_final[[v]],   na.rm = TRUE)), paste0(vars_all,"_sd")))
  )

cluster_desc_with_total <- bind_rows(cluster_desc, total_row)

# 6) Chart-ready means for EXP items (your bar chart style)
exp_cluster_means <- df_final %>%
  group_by(cluster) %>%
  summarise(across(all_of(exp_items), ~mean(., na.rm = TRUE)), .groups = "drop") %>%
  pivot_longer(-cluster, names_to = "exp_item", values_to = "mean_cluster")

exp_total_means <- df_final %>%
  summarise(across(all_of(exp_items), ~mean(., na.rm = TRUE))) %>%
  pivot_longer(everything(), names_to = "exp_item", values_to = "mean_total")

exp_chart_table <- left_join(exp_cluster_means, exp_total_means, by = "exp_item") %>%
  arrange(exp_item, cluster)

# 7) Print the key outputs
print(overall_desc)            # TOTAL n/mean/SD for each variable
print(cluster_desc_with_total) # Per cluster + TOTAL descriptives
print(results_scores)          # Tests for BenLLM_Score & ConLLM_Score
print(results_exp)             # Tests for each EXP item
print(exp_chart_table)         # Means for plotting (cluster vs total) per EXP item

# 8) OPTIONAL: Export everything to Excel
 out_xlsx <- "C:/Users/varis/OneDrive/Desktop/Exeter file/Term 3/R generated file/Significant test vs Average.xlsx"
 wb <- createWorkbook()
addWorksheet(wb, "Overall_Desc");            writeData(wb, "Overall_Desc", overall_desc)
addWorksheet(wb, "Cluster_Desc_TOTAL");      writeData(wb, "Cluster_Desc_TOTAL", cluster_desc_with_total)
addWorksheet(wb, "Tests_SCORES");            writeData(wb, "Tests_SCORES", results_scores)
addWorksheet(wb, "Tests_EXP_Items");         writeData(wb, "Tests_EXP_Items", results_exp)
addWorksheet(wb, "EXP_Chart_Table");         writeData(wb, "EXP_Chart_Table", exp_chart_table)
saveWorkbook(wb, out_xlsx, overwrite = TRUE)
cat(paste0("??? Exported: ", out_xlsx, "\n"))
