# Load necessary libraries
library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(knitr)
library(janitor)
library(gridExtra)
library(grid)
library(writexl)

rm(list = ls())

# Helper function to prettify column names for table titles
prettify_colnames <- function(df) {
  new_names <- names(df) %>%
    gsub("_", " ", .) %>%
    tools::toTitleCase()
  names(df) <- new_names
  df
}

# Helper function to save a data.frame/tibble as a PNG table with prettified column titles
save_table_png <- function(df, filename, width = 2000, height = 1000, res = 300) {
  df_pretty <- prettify_colnames(df)
  tbl <- tableGrob(df_pretty, rows = NULL)
  png(filename, width = width, height = height, res = res)
  grid.draw(tbl)
  dev.off()
}

# Read the Excel file
file_path <- "MTE_VBP Assignement (2022-03-18).xlsx"
excel_sheets(file_path)
data <- read_excel(file_path, sheet = "Price Proposals submitted") %>% 
  clean_names()

data %>% colnames() # display column names

# 1) List unique implant product groups
unique_groups <- unique(data$implant_product_group)
cat("Unique Implant Product Groups:\n")
print(unique_groups)

# 2) Calculate total market revenue for 2020
# a)
revenue_data <- data %>%
  mutate(revenue_2020 = unit_price_excluding_gst_sgd_in_year_2020 * number_of_units_of_each_device_sold_to_all_public_healthcare_institutions_in_year_2020) %>% 
  group_by(implant_product_group, device_company) %>%
  summarise(total_revenue = sum(revenue_2020, na.rm = TRUE), .groups = "drop") %>%
  pivot_wider(names_from = device_company, values_from = total_revenue) %>%
  janitor::adorn_totals(name = "TOTAL")

# b)
revenue_data_bycompany <- revenue_data %>% 
  select(-implant_product_group) %>%
  summarise(across(everything(), sum, na.rm = TRUE))

# 3) Cost effectiveness analysis
cost_effectiveness <- data %>%
  group_by(implant_product_group, device_company) %>%
  summarise(averge_price_2020 = mean(unit_price_excluding_gst_sgd_in_year_2020, na.rm = TRUE)) %>%
  pivot_wider(names_from = device_company, values_from = averge_price_2020) %>%
  clean_names() %>% 
  mutate(more_cost_effective = ifelse(galactic_inc < rocket_company, "Galactic Inc", "Rocket Company"))

# 4a. Average price difference for subsidy
price_diff <- data %>%
  mutate(
    price_diff = unit_price_excluding_gst_sgd_offered_for_subsidy_listing - 
      unit_price_excluding_gst_sgd_in_year_2020
  ) %>%
  group_by(implant_product_group, device_company) %>%
  summarise(ave_diff = mean(price_diff, na.rm = TRUE)) %>%
  pivot_wider(names_from = device_company, values_from = ave_diff)

# 4b) Better prices for subsidy
better_prices <- price_diff %>%
  mutate(
    better_company = case_when(
      `Galactic Inc` < `Rocket Company` ~ "Galactic Inc",
      `Rocket Company` < `Galactic Inc` ~ "Rocket Company",
      TRUE ~ "Equal"
    )
  )

# 5. Recommendation for subsidy
subsidy_recommendation <- data %>%
  group_by(implant_product_group, device_company) %>%
  summarise(avg_subsidy_price = mean(unit_price_excluding_gst_sgd_offered_for_subsidy_listing, na.rm = TRUE)) %>%
  group_by(implant_product_group) %>%
  slice_min(avg_subsidy_price, n = 1) %>%
  select(product_group = implant_product_group, recommended_company = device_company)

# 6. CRT-P price threshold
crtp_threshold <- data %>%
  filter(implant_product_group == "CRT-P") %>%
  summarise(
    Threshold = quantile(unit_price_excluding_gst_sgd_offered_for_subsidy_listing, 0.75),
    .groups = "drop"
  )

# 7. Devices for further price reductions
device_reduction <- data %>%
  mutate(
    price_reduction_pct = (unit_price_excluding_gst_sgd_in_year_2020 - 
                             unit_price_excluding_gst_sgd_offered_for_subsidy_listing) /
      unit_price_excluding_gst_sgd_in_year_2020
  ) %>%
  group_by(device_company) %>%
  arrange(desc(price_reduction_pct)) %>%
  slice_tail(n = 3) %>%  # Select devices with smallest reductions
  select(
    company = device_company,
    device = cardiac_device_name,
    product_Group = implant_product_group,
    subsidy_Price = unit_price_excluding_gst_sgd_offered_for_subsidy_listing,
    price_reduction_pct
  )

# View results (optional)
list(
  Unique_Groups = unique_groups,
  Revenue_Summary = revenue_data,
  Revenue_Summary_ByCompany = revenue_data_bycompany,
  Cost_Effectiveness = cost_effectiveness,
  Price_Difference = price_diff,
  Better_Prices = better_prices,
  Subsidy_Recommendation = subsidy_recommendation,
  CRT_P_Threshold = crtp_threshold,
  Device_Reduction = device_reduction
) %>% print()

# Export to Excel
write_xlsx(
  list(
    "1_Unique_Groups" = data.frame(Groups = unique_groups),
    "2a_Revenue_Summary_Byproduct" = revenue_data,
    '2b_Revenue_Summary_ByCompany' = revenue_data_bycompany,
    "3_Cost_Effectiveness" = cost_effectiveness,
    "4a_Price_Diff" = price_diff,
    "4b_Better_Prices" = better_prices,
    "5_Subsidy_Rec" = subsidy_recommendation,
    "6_CRT_P_Threshold" = crtp_threshold,
    "7_Device_Reductions" = device_reduction
  ),
  "ACE_Analysis_Results.xlsx"
)
# Directory to save PNG tables
png_dir <- "PNG_Tables"

# Create the directory if it doesn't exist
if (!dir.exists(png_dir)) {
  dir.create(png_dir)
}

# Save tables as PNG images with prettified titles in the folder
save_table_png(data.frame(Groups = unique_groups), file.path(png_dir, "1_Unique_Groups.png"))
save_table_png(revenue_data, file.path(png_dir, "2a_Revenue_Summary_Byproduct.png"))
save_table_png(revenue_data_bycompany, file.path(png_dir, "2b_Revenue_Summary_ByCompany.png"))
save_table_png(cost_effectiveness, file.path(png_dir, "3_Cost_Effectiveness.png"))
save_table_png(price_diff, file.path(png_dir, "4a_Price_Diff.png"))
save_table_png(better_prices, file.path(png_dir, "4b_Better_Prices.png"))
save_table_png(subsidy_recommendation, file.path(png_dir, "5_Subsidy_Rec.png"))
save_table_png(crtp_threshold, file.path(png_dir, "6_CRT_P_Threshold.png"))
save_table_png(device_reduction, file.path(png_dir, "7_Device_Reductions.png"), width = 3000, height = 1500)

