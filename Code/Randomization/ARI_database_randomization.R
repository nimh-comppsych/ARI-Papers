# Load necessary packages
library(readxl)
library(ggplot2)
library(openxlsx)
library(dplyr)

# Load excel database
ari_database <- read_xlsx("W:\\string-mbd\\RA Instruction Manuals\\Lily Eisner\\Research\\ARI\\ARI-Papers\\All ARI Papers\\Affective_Reactivity_Index_Papers.xlsx")

# Randomly Select 25 papers to validate
rand_ari_database <- ari_database[sample(nrow(ari_database), 25), ]

# Remove values in Type column
for (row in 1:nrow(rand_ari_database)) {
  rand_ari_database[row,'Type_Primary'] <- NA
  rand_ari_database[row, 'Type_Secondary'] <- NA
}

# Write out random papers to excel
wb <- createWorkbook()
addWorksheet(wb, "Random_ARI")
writeData(wb,"Random_ARI", rand_ari_database)
saveWorkbook(wb, "W:\\string-mbd\\RA Instruction Manuals\\Lily Eisner\\Research\\ARI\\Random_ARI.xlsx", overwrite = TRUE)
