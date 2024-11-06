# --------------------------------------------------
# Title: Descriptive table function
# Description: This function creates a descriptive table as data.frame and save it in a word table
# Developed by: Georges Khoury
# Date: 06/09/2024
# --------------------------------------------------

descriptive_table <- function(data, list = NULL, word = NULL, strat = NULL) {

  # Check for required packages and install if missing
  required_packages <- c("dplyr", "flextable", "officer", "tidyr")

  for (pkg in required_packages) {
    if (!require(pkg, character.only = TRUE)) {
      install.packages(pkg, dependencies = TRUE)
      library(pkg, character.only = TRUE)
    }
  }

  # Function to calculate Mean (IQR) for continuous variables
  fonct_cont = function(data) {
    data1 = data %>% select_if(is.numeric)

    pseudo_percentiles <- function(data, percentiles = c(0.05, 0.5, 0.95)) {
      data <- sort(data)
      qnt <- quantile(data, percentiles)
      qnt_indices <- sapply(qnt, function(q) which.min(abs(data - q)))
      qnt_indices[qnt_indices < 3] <- 3
      qnt_indices[qnt_indices > length(data) + 2] <- length(data) - 2
      qnt_means <- sapply(qnt_indices, function(i) mean(data[max(1, i-2) : min(length(data), i+2)]))
      return(qnt_means)
    }

    result <- apply(data1, 2, pseudo_percentiles)
    result_data <- as.data.frame(t(result))

    mean = result_data$`50%`
    IQR = result_data$`95%` - result_data$`5%`

    mean_IQR = paste0(sprintf("%.1f", mean), " (", sprintf("%.1f", IQR), ")")

    na_count = sapply(data1, function(x) sum(is.na(x)))
    n_count = sapply(data1, function(x) sum(!is.na(x)))

    n_info = as.character(n_count)

    final_table <- data.frame(Variable = character(0),
                              Total_N = character(0),
                              Levels = character(0),
                              Stat = character(0),
                              stringsAsFactors = FALSE)

    for (i in seq_along(colnames(data1))) {
      final_table[nrow(final_table) + 1, ] <- list(
        Variable = colnames(data1)[i],
        Total_N = n_info[i],
        Levels = "Median(Range)",
        Stat = mean_IQR[i]
      )

      if (na_count[i] > 0) {
        final_table[nrow(final_table) + 1, ] <- list(
          Variable = colnames(data1)[i],
          Total_N = "",
          Levels = "NA",
          Stat = as.character(na_count[i])
        )
      }
    }

    return(final_table)
  }

  # Function to calculate percentages for categorical variables
  fonct_categorical = function(data) {
    data2 = data %>% select_if(Negate(is.numeric))
    final_table = data.frame()

    for (col_name in colnames(data2)) {
      column_data = data2[[col_name]]
      total_obs = sum(!is.na(column_data))
      n_obs = table(column_data, useNA = "no")
      percentages = (n_obs / total_obs) * 100

      result_data = data.frame(
        Variable = rep(col_name, length(n_obs)),
        Levels = names(n_obs),
        Total_N = sprintf("%d", n_obs),
        Stat = paste0(sprintf("%d", n_obs), " (", sprintf("%.1f", percentages), "%)"),
        stringsAsFactors = FALSE
      )

      na_count = sum(is.na(column_data))
      if (na_count > 0) {
        na_info = paste0(sprintf("%d", na_count), " (", sprintf("%.1f", (na_count / total_obs) * 100), "%)")
        na_row = data.frame(
          Variable = col_name,
          Total_N = "",
          Levels = "NA",
          Stat = na_info,
          stringsAsFactors = FALSE
        )
      } else {
        na_row = NULL
      }

      final_table = rbind(final_table, result_data)

      if (!is.null(na_row)) {
        final_table = rbind(final_table, na_row)
      }
    }

    rownames(final_table) = NULL
    final_table
  }

  # Calculate the continuous and categorical results
  continuous_results = fonct_cont(data)
  categorical_results = fonct_categorical(data)

  combined_results <- bind_rows(continuous_results, categorical_results)

  if (!is.null(list)) {
    combined_results <- combined_results %>%
      filter(Variable %in% list) %>%
      arrange(match(Variable, list))
  }

  if (!is.null(strat)) {
    strat_levels <- unique(data[[strat]])
    strat_results <- lapply(strat_levels, function(level) {
      strat_data <- data %>% filter(!!sym(strat) == !!level)
      continuous_strat <- fonct_cont(strat_data)
      categorical_strat <- fonct_categorical(strat_data)
      combined_strat <- bind_rows(continuous_strat, categorical_strat)
      combined_strat <- combined_strat %>%
        mutate(Stratum = level)
      combined_strat
    })

    # Combine results side by side, aligning by Variable and Levels
    strat_results_df <- Reduce(function(x, y) {
      # Ensure columns to rename exist
      rename_and_merge <- function(df, stratum) {
        cols_to_rename <- intersect(c("Total_N", "Stat"), names(df))
        if (length(cols_to_rename) > 0) {
          df <- df %>%
            rename_with(~ paste0(.x, "_", stratum), .cols = all_of(cols_to_rename)) %>%
            select(-Stratum)  # Remove the Stratum column after renaming
        }
        df
      }

      x <- rename_and_merge(x, unique(x$Stratum))
      y <- rename_and_merge(y, unique(y$Stratum))

      full_join(x, y, by = c("Variable", "Levels"))
    }, strat_results)

    # Remove the stratification variable itself from the results
    final_results <- strat_results_df %>%
      filter(Variable != strat) %>%
      rename_with(~ gsub("Total_N", "N", .), .cols = starts_with("Total_N"))

    # Apply the list filtering after stratification
    if (!is.null(list)) {
      final_results <- final_results %>%
        filter(Variable %in% list) %>%
        arrange(match(Variable, list))
    }
  } else {
    final_results <- combined_results

    # Apply the list filtering
    if (!is.null(list)) {
      final_results <- final_results %>%
        filter(Variable %in% list) %>%
        arrange(match(Variable, list))
    }
  }

  # Reorder the columns: Levels before Total_N or N
  final_results <- final_results %>%
    select(Variable, Levels, everything())

  # Prepare the final formatted results for the Word document
  formatted_results <- final_results %>%
    mutate(Variable = ifelse(duplicated(Variable), "", Variable))

  if (!is.null(word)) {
    formatted_results$Variable <- sub("^([a-z])", "\\U\\1", formatted_results$Variable, perl = TRUE)
  }

  ft <- flextable(formatted_results) %>%
    set_table_properties(width = 1, layout = "autofit") %>%
    theme_vanilla() %>%
    autofit() %>%
    align(align = "center", part = "all") %>%
    bold(part = "header") %>%
    fontsize(size = 10, part = "all") %>%
    set_header_labels(
      Variable = "Variable",
      Levels = "Levels",
      Total_N = ifelse(is.null(strat), "Total N", "N"),
      Stat = "Stat¹"
    ) %>%
    bold(j = "Variable", bold = TRUE, part = "body") %>%
    add_footer_lines(values = "¹ Stat: Median [Range] for continuous variables, N (%) for categorical/binary variables.") %>%
    add_footer_lines(values = "Median and percentiles were based on information from at least five individuals with values closest to the actual median/percentile.") %>%
    fontsize(size = 8, part = "footer")

  ft <- border_inner_h(ft, border = fp_border(width = 0), part = "body")

  if (!is.null(word)) {
    doc <- read_docx() %>%
      body_add_flextable(value = ft)
    print(doc, target = word)
  }

  return(final_results)
}
