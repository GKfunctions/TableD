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



# --------------------------------------------------
# Title: Correlation matrix function
# Description: This function creates a correlation matrix
# Developed by: Georges Khoury
# Date: 06/09/2024
# --------------------------------------------------

correlation <- function(data, list = NULL, word = NULL, pdf = NULL, title = NULL) {

  # Check for required packages and install if missing
  required_packages <- c("dplyr", "corrplot", "officer", "ggplot2")

  for (pkg in required_packages) {
    if (!require(pkg, character.only = TRUE)) {
      install.packages(pkg, dependencies = TRUE)
      library(pkg, character.only = TRUE)
    }
  }

  # Check if the list argument is provided
  if (is.null(list)) {
    stop("Please provide a list of variables.")
  }

  # Calculate the correlation matrix
  corr_matrix <- cor(data[list], method = "spearman", use = "pairwise.complete.obs")

  # Color palette for the plot
  color_palette <- colorRampPalette(c("dodgerblue", "white", "brown1"))(300)

  # Display the plot in R
  corrplot(
    corr_matrix,
    method = "color",
    type = "lower",
    col = color_palette,
    tl.pos = "tl",
    addCoef.col = "black",
    number.cex = 0.8,
    tl.cex = 0.9,
    tl.col = "black",
    cl.pos = "r",
    tl.srt = 45
  )

  # Save the plot as PDF if 'pdf' argument is provided
  if (!is.null(pdf)) {
    # Save plot as PDF
    pdf(pdf, width = 6, height = 6)
    corrplot(
      corr_matrix,
      method = "color",
      type = "lower",
      col = color_palette,
      tl.pos = "tl",
      addCoef.col = "black",
      number.cex = 0.8,
      tl.cex = 0.9,
      tl.col = "black",
      cl.pos = "r",
      tl.srt = 45
    )
    dev.off()
    message("Plot saved as PDF: ", pdf)
  }

  # Save the plot to a Word file if 'word' argument is provided
  if (!is.null(word)) {
    # Create a temporary file for the plot (as PNG for Word)
    temp_plot <- tempfile(fileext = ".png")

    # Save the corrplot as a PNG file for Word
    png(temp_plot, width = 6, height = 6, units = "in", res = 300)
    corrplot(
      corr_matrix,
      method = "color",
      type = "lower",
      col = color_palette,
      tl.pos = "tl",
      addCoef.col = "black",
      number.cex = 0.8,
      tl.cex = 0.9,
      tl.col = "black",
      cl.pos = "r",
      tl.srt = 45
    )
    dev.off()  # Close the graphical device

    # Create a Word document object
    doc <- read_docx()

    # Add the PNG plot to the Word document
    doc <- doc %>%
      body_add_img(src = temp_plot, width = 6, height = 6)

    # If a title is provided, add it after the plot
    if (!is.null(title)) {
      # Define text formatting properties (Times New Roman, size 12)
      title_format <- fp_text(font.size = 12, font.family = "Times New Roman")

      # Add the title paragraph with the specified formatting
      doc <- doc %>%
        body_add_fpar(fpar(ftext(title, title_format)))
    }

    # Save the Word document to the specified file location
    print(doc, target = word)
    message("Plot and title saved to Word file: ", word)
  }
}



# --------------------------------------------------
# Title: Variable decomposition function
# Description: This function decomposes variables into quartile specified between 2 and 4 and saves a word document with the percentiles calculated according to the pseudo percentile
# Developed by: Georges Khoury
# Date: 06/09/2024
# --------------------------------------------------

decomp <- function(data, list = NULL, q = NULL, suffix = NULL, word = NULL) {

  required_packages <- c("dplyr", "flextable", "officer", "tidyr")

  for (pkg in required_packages) {
    if (!require(pkg, character.only = TRUE)) {
      install.packages(pkg, dependencies = TRUE)
      library(pkg, character.only = TRUE)
    }
  }

  pseudo_percentiles <- function(data, percentiles = c(0.05, 0.25, 0.33, 0.5, 0.67, 0.75, 0.95)) {
    data <- sort(data)
    qnt <- quantile(data, percentiles)
    qnt_indices <- sapply(qnt, function(q) which.min(abs(data - q)))
    qnt_indices[qnt_indices < 3] <- 3
    qnt_indices[qnt_indices > length(data) + 2] <- length(data) - 2
    qnt_means <- sapply(qnt_indices, function(i) mean(data[max(1, i-2):min(length(data), i+2)]))
    names(qnt_means) <- paste0("P", gsub("%", "", names(qnt)))
    return(qnt_means)
  }

  # Check if 'list', 'q', 'suffix', and 'word' are provided
  if (is.null(list) || is.null(q) || is.null(suffix) || is.null(word)) {
    stop("Error: 'list', 'q', 'suffix', and 'word' parameters must be provided.")
  }

  # Check if 'q' is between 2 and 4
  if (q < 2 || q > 4) {
    stop("Error: 'q' must be between 2 and 4.")
  }

  # Prepare a Word document
  doc <- read_docx()

  # Define the percentiles based on 'q'
  percentile_list <- switch(as.character(q),
                            "2" = c(0.05, 0.5, 0.95),
                            "3" = c(0.05, 0.33, 0.67, 0.95),
                            "4" = c(0.05, 0.25, 0.5, 0.75, 0.95))

  # Iterate over the provided list of variables
  for (var in list) {

    # Check if the variable exists in the data
    if (!var %in% colnames(data)) {
      stop(paste("Error: Variable", var, "not found in the data."))
    }

    # Create a new column name by adding the suffix to the variable
    new_col <- paste(var, suffix, sep = "")

    # Check if the new variable already exists in the data (to prevent overwriting)
    if (new_col %in% colnames(data)) {
      warning(paste("Warning: Variable", new_col, "already exists in the data. Overwriting it."))
    }

    # Apply the quantile cut to the variable and create the new column
    data[[new_col]] <- cut(data[[var]],
                           breaks = quantile(data[[var]], probs = seq(0, 1, length.out = q + 1), na.rm = TRUE),
                           include.lowest = TRUE,
                           labels = FALSE)

    # Get the pseudo percentiles for the variable
    percentiles <- pseudo_percentiles(data[[var]], percentiles = percentile_list)

    # Prepare the percentile table for Word
    table_data <- data.frame(
      Statistic = c("Min", names(percentiles), "Max"),
      Value = c(min(data[[var]], na.rm = TRUE), percentiles, max(data[[var]], na.rm = TRUE))
    )

    # Add a title for the variable
    doc <- body_add_par(doc, paste("Percentile Information for", var), style = "heading 2")

    # Add the table to the Word document (without style argument)
    doc <- body_add_table(doc, value = table_data)
  }

  # Add footer
  doc <- body_add_par(doc, "", style = "Normal")
  doc <- body_add_par(doc, "Percentiles were based on information from at least five individuals with values closest to the actual median/percentile.", style = "Normal")

  # Save the Word document to the provided location
  print(doc, target = word)

  return(data)
}



# --------------------------------------------------
# Title: Limit of detection function
# Description: This function creates a word document for the limit of detection (LOD)
# Developed by: Georges Khoury
# Date: 06/09/2024
# --------------------------------------------------


LOD <- function(data, list = NULL, word = NULL) {

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
    # If a specific list of variables is provided, subset the data accordingly
    if (!is.null(list)) {
      data <- data %>% select(all_of(list))
    }

    # Select only numeric columns
    data1 = data %>% select_if(is.numeric)

    # Pseudo percentiles calculation function
    pseudo_percentiles <- function(data, percentiles = c(0.05, 0.25, 0.5, 0.75, 0.95)) {
      data <- sort(data)
      qnt <- quantile(data, percentiles)
      qnt_indices <- sapply(qnt, function(q) which.min(abs(data - q)))
      qnt_indices[qnt_indices < 3] <- 3
      qnt_indices[qnt_indices > length(data) + 2] <- length(data) - 2
      qnt_means <- sapply(qnt_indices, function(i) mean(data[max(1, i-2) : min(length(data), i+2)]))
      return(qnt_means)
    }

    # Apply the pseudo_percentiles function to each numeric column
    result <- apply(data1, 2, pseudo_percentiles)
    result_data <- as.data.frame(t(result))

    # Naming the columns appropriately
    colnames(result_data) <- c("5%", "25%", "50%", "75%", "95%")

    # Add variable names as a column
    result_data <- result_data %>%
      mutate(Variable = rownames(result_data)) %>%
      select(Variable, everything())

    result_data <- result_data %>%
      mutate(across(c("5%", "25%", "50%", "75%", "95%"), ~ format(round(as.numeric(.), 2), nsmall = 2)))

    # Add the empty columns
    result_data <- result_data %>%
      mutate(`Batch 1` = "",
             `Batch 2` = "",
             `N (%) > LOD` ="" )%>%
      select(Variable, `Batch 1`, `Batch 2`, `N (%) > LOD`, everything())

    return(result_data)
  }

  # Call the continuous variables function and get the resulting table
  cont_table <- fonct_cont(data)

  # Create the flextable
  ft <- flextable(cont_table)

  # Modify the flextable to use Times New Roman and size 12
  ft <- ft %>%
    font(fontname = "Times New Roman", part = "all") %>%
    fontsize(size = 12, part = "all") %>%
    add_footer_lines(values = "Median and percentiles were based on information from at least five individuals with values closest to the actual median/percentile.") %>%
    autofit()

  # Apply superscript for the percentile column headers
  ft <- set_header_labels(ft, `5%` = "5P", `25%` = "25P", `50%` = "50P", `75%` = "75P", `95%` = "95P")
  ft <- compose(ft, j = "5%", part = "header", value = as_paragraph("5", as_sup("th"), " P"))
  ft <- compose(ft, j = "25%", part = "header", value = as_paragraph("25", as_sup("th"), " P"))
  ft <- compose(ft, j = "50%", part = "header", value = as_paragraph("50", as_sup("th"), " P"))
  ft <- compose(ft, j = "75%", part = "header", value = as_paragraph("75", as_sup("th"), " P"))
  ft <- compose(ft, j = "95%", part = "header", value = as_paragraph("95", as_sup("th"), " P"))

  # If a file location is provided in 'word', save the flextable to a Word file
  if (!is.null(word)) {
    doc <- read_docx() %>%
      body_add_flextable(ft) %>%
      body_add_par(" ", style = "Normal")

    # Save the Word document
    print(doc, target = word)
  }

  # Return the result as a data.frame
  return(cont_table)
}
