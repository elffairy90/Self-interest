library(readxl)
library(dplyr)
library(tidyr)
library(ggplot2)
library(knitr)
library(janitor)
library(gridExtra)
library(grid)
library(writexl)
library(scales)
library(stringr)
library(patchwork) 



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

# load files
file_path <- "MTE_VBP Assignement (2022-03-18).xlsx"
excel_sheets(file_path)
data <- read_excel(file_path, sheet = "Price Proposals submitted") %>% 
  clean_names()


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

revenue_data_dollar <- revenue_data %>% 
  mutate(across(-implant_product_group, ~ dollar(.x, prefix = "S$ ", accuracy = 0.01)))


# # b) can extract from a)


# 3) Cost effectiveness analysis
cost_effectiveness <- data %>%
  group_by(implant_product_group, device_company) %>%
  summarise(
    average_price_2020 = mean(unit_price_excluding_gst_sgd_in_year_2020, na.rm = TRUE),
    total_units_sold = sum(number_of_units_of_each_device_sold_to_all_public_healthcare_institutions_in_year_2020, na.rm = TRUE),
    total_cost = sum(unit_price_excluding_gst_sgd_in_year_2020 * number_of_units_of_each_device_sold_to_all_public_healthcare_institutions_in_year_2020, na.rm = TRUE)
  ) %>%
  pivot_wider(
    names_from = device_company,
    values_from = c(average_price_2020, total_units_sold, total_cost)
  ) %>%
  clean_names() %>%
  mutate(
    more_cost_effective_by_price = ifelse(average_price_2020_galactic_inc < average_price_2020_rocket_company, "Galactic Inc", "Rocket Company"),
    more_cost_effective_by_total_cost = ifelse(total_cost_galactic_inc < total_cost_rocket_company, "Galactic Inc", "Rocket Company")
  ) %>%
  mutate(across(
    starts_with("average_price_2020_"),
    ~ dollar(.x, prefix = "S$ ", accuracy = 0.01)
  )) %>%
  mutate(across(
    starts_with("total_cost_"),
    ~ dollar(.x, prefix = "S$ ", accuracy = 0.01)
  ))



# 4a. Average price difference for subsidy
price_diff_clean <- data %>%
  group_by(implant_product_group, device_company) %>%
  summarise(
    avg_price_2020 = mean(unit_price_excluding_gst_sgd_in_year_2020, na.rm = TRUE),
    avg_price_subsidy = mean(unit_price_excluding_gst_sgd_offered_for_subsidy_listing, na.rm = TRUE),
    .groups = "drop"
  ) %>%
  pivot_wider(names_from = device_company, values_from = c(avg_price_2020, avg_price_subsidy)) %>%
  clean_names() %>% 
  mutate(
    diff_price_2020 = avg_price_2020_galactic_inc - avg_price_2020_rocket_company,
    diff_price_subsidy = avg_price_subsidy_galactic_inc - avg_price_subsidy_rocket_company
  ) 

price_diff <- price_diff_clean %>%
  mutate(across(starts_with("avg_price"), ~ dollar(.x, prefix = "S$ ", accuracy = 0.01)),
         across(starts_with("diff_price"), ~ dollar(.x, prefix = "S$ ", accuracy = 0.01))) 

price_diff %>% head()

# 4b) Better prices for subsidy
better_prices <- price_diff_clean %>%
  mutate(
    better_price = ifelse(
      avg_price_subsidy_galactic_inc < avg_price_subsidy_rocket_company,
      "Galactic Inc",
      "Rocket Company"
    )
  ) %>% 
  select (implant_product_group, avg_price_subsidy_galactic_inc, avg_price_subsidy_rocket_company, better_price) %>% 
  mutate(across(-c(implant_product_group, better_price), ~ dollar(.x, prefix = "S$ ", accuracy = 0.01)))



# 5. The better compnay to recommendate 
summary_table <- data %>%
  group_by(implant_product_group, device_company) %>%
  summarise(
    avg_price = mean(unit_price_excluding_gst_sgd_in_year_2020, na.rm = TRUE),
    avg_subsidy_price = mean(unit_price_excluding_gst_sgd_offered_for_subsidy_listing, na.rm = TRUE),
    avg_net_price = avg_subsidy_price,  
    total_volume = sum(number_of_units_of_each_device_sold_to_all_public_healthcare_institutions_in_year_2020, na.rm = TRUE),
    avg_longevity = mean(stated_longevity_years, na.rm = TRUE),
    avg_warranty = mean(device_warranty_years, na.rm = TRUE),
    mri_compatibility_pct = mean(mri_compatibility_y_n == "Y", na.rm = TRUE),   # % with MRI compatibility
    remote_monitoring_pct = mean(home_remote_monitoring_system_compatibility_y_n == "Y", na.rm = TRUE),
    hsa_registered_pct = mean(hsa_registration_status_registered_pending_approval_in_the_pipeline == "Registered", na.rm = TRUE)
  ) %>%
  ungroup()

ranked_table <- summary_table %>%
  group_by(implant_product_group) %>%
  mutate(
    rank_price = rank(avg_price, ties.method = "min"),
    rank_net_price = rank(avg_net_price, ties.method = "min"),
    rank_volume = rank(desc(total_volume), ties.method = "min"),
    rank_longevity = rank(desc(avg_longevity), ties.method = "min"),
    rank_warranty = rank(desc(avg_warranty), ties.method = "min"),
    rank_mri_compatibility = rank(desc(mri_compatibility_pct), ties.method = "min"),
    rank_remote_monitoring = rank(desc(remote_monitoring_pct), ties.method = "min"),
    rank_hsa_registered = rank(desc(hsa_registered_pct), ties.method = "min"),
    
    combined_score = rank_price + rank_net_price + rank_volume + rank_longevity +
      rank_warranty + rank_mri_compatibility + rank_remote_monitoring + rank_hsa_registered
  ) %>%
  ungroup()


subsidy_recommendation <- ranked_table %>%
  group_by(implant_product_group) %>%
  slice_min(combined_score, n = 1) %>%
  select(
    implant_product_group, device_company, avg_price, avg_net_price, total_volume,
    avg_longevity, avg_warranty, mri_compatibility_pct, remote_monitoring_pct, hsa_registered_pct,
    combined_score
  )

# -------------------------------------------------
################ 6.CRTP Threshold #################
# -------------------------------------------------
# 1. Data Preparation 
crt_p_data <- data %>%
  filter(implant_product_group == "CRT-P") %>%
  # Clean and transform key columns
  mutate(
    subsidy_price = as.numeric(unit_price_excluding_gst_sgd_offered_for_subsidy_listing),
    longevity = as.numeric(stated_longevity_years),
    units_sold = as.numeric(number_of_units_of_each_device_sold_to_all_public_healthcare_institutions_in_year_2020),
    # Create feature tiers
    feature_tier = case_when(
      mri_compatibility_y_n == "Y" & home_remote_monitoring_system_compatibility_y_n == "Y" ~ "Premium",
      mri_compatibility_y_n == "Y" | home_remote_monitoring_system_compatibility_y_n == "Y" ~ "Mid-Tier",
      TRUE ~ "Basic"
    )
  ) %>%
  filter(!is.na(subsidy_price))  # Remove devices without subsidy pricing

# 2. Price Distribution Analysis ------------------------------------------
# Calculate weighted statistics (volume-adjusted)
price_summary <- crt_p_data %>%
  summarise(
    min_price = min(subsidy_price),
    q1_price = quantile(subsidy_price, 0.25),
    median_price = median(subsidy_price),
    mean_price = mean(subsidy_price),
    q3_price = quantile(subsidy_price, 0.75),
    max_price = max(subsidy_price),
    # Volume-weighted mean
    vol_weighted_mean = weighted.mean(subsidy_price, w = units_sold, na.rm = TRUE),
    # Volume at median price
    vol_at_median = sum(units_sold[subsidy_price <= median(subsidy_price)])
  ) %>%
  mutate(across(where(is.numeric), ~ round(.x, 2)))

# 3. Tiered Threshold Calculation -----------------------------------------
tiered_thresholds <- crt_p_data %>%
  group_by(feature_tier) %>%
  summarise(
    devices = n(),
    units = sum(units_sold, na.rm = TRUE),
    tier_threshold = median(subsidy_price),
    .groups = "drop"
  ) %>%
  mutate(
    market_coverage = units / sum(units),
    # Format for reporting
    tier_threshold_fmt = dollar(tier_threshold, prefix = "S$ ", accuracy = 0.01),
    market_coverage_pct = percent(market_coverage, accuracy = 0.1)
  )

# 4. Savings Projection ---------------------------------------------------
crt_p_data_with_threshold <- crt_p_data %>%
  left_join(
    tiered_thresholds %>% select(feature_tier, tier_threshold),
    by = "feature_tier"
  ) %>%
  mutate(
    potential_savings = ifelse(subsidy_price > tier_threshold,
                               (subsidy_price - tier_threshold) * units_sold,
                               0)
  )
savings_breakdown <- crt_p_data_with_threshold %>%
  group_by(feature_tier) %>%
  summarise(
    tier_savings = sum(potential_savings, na.rm = TRUE)
  ) %>%
  pivot_wider(names_from = feature_tier, values_from = tier_savings, names_prefix = "savings_") %>%
  mutate(total_potential_savings = rowSums(across(starts_with("savings_")), na.rm = TRUE)) %>% 
  mutate(across(everything(), ~ dollar(.x, prefix = "S$ ", accuracy = 0.01)))

# 5. Visualization --------------------------------------------------------
# Price distribution plot
p1 <- ggplot(crt_p_data, aes(x = subsidy_price)) +
  geom_histogram(fill = "steelblue", bins = 15, alpha = 0.8) +
  geom_vline(data = tiered_thresholds, 
             aes(xintercept = tier_threshold, color = feature_tier), 
             linetype = "dashed", size = 1.2) +
  scale_x_continuous(labels = dollar_format(prefix = "S$")) +
  labs(
    title = "CRT-P Device Price Distribution with Tiered Thresholds",
    subtitle = "Vertical lines show recommended price thresholds by feature tier",
    x = "Subsidy Price (SGD)",
    y = "Number of Devices"
  ) +
  theme_minimal() +
  theme(legend.position = "top")

# Price vs Longevity by Tier
p2 <- ggplot(crt_p_data, aes(x = longevity, y = subsidy_price, color = feature_tier)) +
  geom_point(size = 3, alpha = 0.7) +
  geom_smooth(method = "lm", se = FALSE, size = 0.8) +
  scale_y_continuous(labels = dollar_format(prefix = "S$")) +
  labs(
    title = "Price vs Longevity Relationship by Feature Tier",
    x = "Stated Longevity (Years)",
    y = "Subsidy Price (SGD)",
    color = "Feature Tier"
  ) +
  theme_minimal()

# Combine plots
combined_plot <- p1 / p2 + plot_layout(heights = c(1, 1))
ggsave("CRT-P_Price_Analysis.png", plot = combined_plot, width = 10, height = 8, dpi = 300)

ggsave("PNG_Tables/Tier.png", plot = p1, width = 10, height = 8, dpi = 300)

# 6. Company-Level Analysis 
company_summary <- crt_p_data %>%
  group_by(device_company) %>%
  summarise(
    devices = n(),
    avg_price = mean(subsidy_price),
    max_price = max(subsidy_price),
    min_price = min(subsidy_price),
    units_sold = sum(units_sold),
    .groups = "drop"
  ) %>%
  mutate(
    price_status = case_when(
      max_price > tiered_thresholds$tier_threshold[3] ~ "Above Premium Threshold",
      max_price > tiered_thresholds$tier_threshold[2] ~ "Above Mid-Tier Threshold",
      TRUE ~ "Within All Thresholds"
    ),
    # Formatting
    avg_price_fmt = dollar(avg_price, prefix = "S$ ", accuracy = 0.01),
    max_price_fmt = dollar(max_price, prefix = "S$ ", accuracy = 0.01),
    min_price_fmt = dollar(min_price, prefix = "S$ ", accuracy = 0.01)
  ) %>%
  arrange(desc(avg_price))


# Final Reporting Output
cat("CRT-P SUBSIDY PRICE THRESHOLD RECOMMENDATION\n")
cat("============================================\n\n")
cat("Based on comprehensive analysis of", nrow(crt_p_data), "CRT-P devices:\n\n")

# Print tiered thresholds table
cat("Recommended Tiered Thresholds:\n")
cat("------------------------------\n")
tiered_thresholds %>%
  select(Feature_Tier = feature_tier,
         Threshold = tier_threshold_fmt,
         Devices = devices,
         Units = units,
         Market_Coverage = market_coverage_pct) %>%
  print()


# Key Insights
cat("Key Insights:\n")
cat("-------------\n")

premium_pct <- mean(crt_p_data$feature_tier == "Premium", na.rm = TRUE)

cat("- ", percent(premium_pct, accuracy = 0.1), " of devices qualify as Premium tier.\n")
cat("- Thresholds cover ", percent(sum(tiered_thresholds$market_coverage), accuracy = 0.1), " of current market volume.\n")

affected_models <- crt_p_data_with_threshold %>%
  filter(subsidy_price > tier_threshold) %>%
  nrow()

cat("- Implementation would affect ", affected_models, " device models currently priced above thresholds.\n")


# Company Impact (Assuming company_summary already exists)
cat("\nCompany Impact Summary:\n")
cat("-----------------------\n")



# -------------------------------------------------
###### 7. Devices for further price reductions ###### 
# -------------------------------------------------
data$cardiac_device_name %>% unique()


# Preprocess data: convert columns, calculate feature score, and filter HSA-registered devices
data_prep <- data %>%
  mutate(
    # Convert key columns to numeric
    volume_2020 = as.numeric(number_of_units_of_each_device_sold_to_all_public_healthcare_institutions_in_year_2020),
    subsidy_price = as.numeric(unit_price_excluding_gst_sgd_offered_for_subsidy_listing),
    original_price = as.numeric(unit_price_excluding_gst_sgd_in_year_2020),
    
    # Calculate feature score (1 point per advanced feature: MRI + Remote Monitoring)
    feature_score = ifelse(mri_compatibility_y_n == "Y", 1, 0) + 
      ifelse(home_remote_monitoring_system_compatibility_y_n == "Y", 1, 0),
    
    # Handle missing values
    volume_2020 = ifelse(is.na(volume_2020), 0, volume_2020)
  ) %>%
  filter(
    # Keep only HSA-approved or pending devices
    hsa_registration_status_registered_pending_approval_in_the_pipeline %in% 
      c("Registered", "Pending Approval", "In the Pipeline"),
    !is.na(subsidy_price),       # Remove devices without subsidy pricing
    !is.na(original_price),      # Remove devices missing original price
    subsidy_price <= original_price  # Exclude devices where subsidy price > original price
  ) %>%
  mutate(
    # Calculate priority score: (Volume × Subsidy Price) × (1 + Feature Bonus)
    priority_score = volume_2020 * subsidy_price * (1 + 0.1 * feature_score),
    
    # Format subsidy_price with dollar sign
    subsidy_price = dollar(subsidy_price, prefix = "S$ ")
  )

# Select top 3 devices per company
shortlist <- data_prep %>%
  group_by(device_company) %>%
  arrange(desc(priority_score), .by_group = TRUE) %>%
  slice_head(n = 3) %>%
  ungroup() %>%
  select(
    device_company,
    cardiac_device_name,
    product_code,
    volume_2020,
    subsidy_price,
    mri_compatibility_y_n,
    home_remote_monitoring_system_compatibility_y_n,
    priority_score
  )

# View results
print(shortlist)

device_reduction_top3 <- shortlist
# View results
print(device_reduction_top3)

#####################  Export to Excel ####################################- 
write_xlsx(
  list(
    "1_Unique_Groups" = data.frame(Groups = unique_groups),
    "2a_Revenue_Summary_Byproduct" = revenue_data,
    # '2b_Revenue_Summary_ByCompany' = revenue_data_bycompany,
    "3_Cost_Effectiveness" = cost_effectiveness,
    "4a_Price_Diff" = price_diff,
    "4b_Better_Prices" = better_prices,
    "5_Subsidy_Rec" = subsidy_recommendation,
    "6_CRT_P_Threshold" = company_summary,
    "6_price_summary" = price_summary,
    "6_tiered_thresholds" = tiered_thresholds,
    # "CRT-P_Price_Analysis" = combined_plot,
    "7_Device_Reductions" = device_reduction_top3
  ),
  "ACE_Analysis_Results.xlsx"
)
# Directory to save PNG tables
png_dir <- "PNG_Tables"

# Create the directory if it doesn't exist
if (!dir.exists(png_dir)) {
  dir.create(png_dir)
}


#----------------------
# Save tables as PNG images with prettified titles in the folder
save_table_png(data.frame(Groups = unique_groups), file.path(png_dir, "1_Unique_Groups.png"))
save_table_png(revenue_data_dollar, file.path(png_dir, "2a_Revenue_Summary_Byproduct.png"))
# save_table_png(revenue_data_bycompany_dollar, file.path(png_dir, "2b_Revenue_Summary_ByCompany.png"))
save_table_png(cost_effectiveness, file.path(png_dir, "3_Cost_Effectiveness.png"), width =9000, height = 1500)
save_table_png(price_diff, file.path(png_dir, "4a_Price_Diff.png"),  width = 6000, height = 1500)
save_table_png(better_prices, file.path(png_dir, "4b_Better_Prices.png"), width = 6000, height = 1500)
save_table_png(subsidy_recommendation, file.path(png_dir, "5_Subsidy_Rec.png"), width = 6000, height = 1500)
save_table_png(price_summary, file.path(png_dir, "6_price_summary.png"), width = 3000, height = 500)
save_table_png(company_summary, file.path(png_dir, "6_CRT_P_Threshold.png"), width = 4000, height = 500)
save_table_png(tiered_thresholds, file.path(png_dir, "6_tiered_thresholds.png"), width = 3000, height = 800)
save_table_png(device_reduction_top3, file.path(png_dir, "7_Device_Reductions.png"), width = 5500, height = 1000)


