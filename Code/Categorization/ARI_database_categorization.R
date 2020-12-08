# Load necessary packages
library(readxl)
library(ggplot2)
library(openxlsx)
library(dplyr)

# Load excel database
ari_database <- read_xlsx("W:\\string-mbd\\RA Instruction Manuals\\Lily Eisner\\Research\\ARI\\ARI-Papers\\All ARI Papers\\Affective_Reactivity_Index_Papers.xlsx")

# Split database by primary categorization in psychometric, observational, and treatment
ari_database_psychometric <- ari_database %>% filter(Type_Primary == "Psychometric")
ari_database_observational <- ari_database %>% filter(Type_Primary == "Observational")
ari_database_treatment <-ari_database %>% filter(Type_Primary == "Treatment")

# Write out psychometric to excel
wb <- createWorkbook()
addWorksheet(wb, "Psychometric_Papers")
writeData(wb,"Psychometric_Papers", ari_database_psychometric)
saveWorkbook(wb, "W:\\string-mbd\\RA Instruction Manuals\\Lily Eisner\\Research\\ARI\\ARI-Papers\\Psychometric Papers\\Psychometric_Papers.xlsx", overwrite = TRUE)

# Write out observational to excel
wb <- createWorkbook()
addWorksheet(wb, "Observational_Papers")
writeData(wb,"Observational_Papers", ari_database_observational)
saveWorkbook(wb, "W:\\string-mbd\\RA Instruction Manuals\\Lily Eisner\\Research\\ARI\\ARI-Papers\\Observational Papers\\Observational_Papers.xlsx", overwrite = TRUE)

# Write out treatment to excel
wb <- createWorkbook()
addWorksheet(wb, "Treatment_Papers")
writeData(wb,"Treatment_Papers", ari_database_treatment)
saveWorkbook(wb, "W:\\string-mbd\\RA Instruction Manuals\\Lily Eisner\\Research\\ARI\\ARI-Papers\\Treatment Papers\\Treatment_Papers.xlsx", overwrite = TRUE)
