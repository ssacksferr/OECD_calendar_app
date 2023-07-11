library(readxl)
library(dplyr)
library(tidyverse)
library(countrycode)

##First, read the excel file that copies COLUMN M of QNA_CALENDAR. Copy it 
## in a sheet and =CLEAN() the contents of COLUMN M so they are all 
##in one line.

##Named the new, clean column "trimmed". 

##Make sure all items are separated by "; ".

qna_cal <- read_excel("C:/Users/sacksferrari_s/OneDrive - OECD/R Projects and Files/calendar_app/calendar_app/import_qna_cal.xlsx")
qna_cal <- qna_cal %>% drop_na(trimmed)


process_var3 <- function(var3) {
  elements <- unlist(strsplit(var3, ";\\s*|;"))
  
  result <- list()
  counter <- 1
  current_Q <- ""
  
  for (element in elements) {
    if (startsWith(element, "Q")) {
      current_Q <- element
      counter <- 1
    } else {
      result <- c(result, paste0(element, " ",counter, " ", current_Q))
      counter <- counter + 1
    }
  }
  
  return(result)
}

df_processed <- qna_cal %>%
  group_by(country) %>%
  mutate(var3_processed = list(process_var3(trimmed))) %>%
  filter(!is.na(var3_processed)) %>%
  tidyr::unnest_wider(var3_processed, names_sep = "_") %>%
  rename_with(~ paste0("var3_", seq_along(.)), starts_with("var3_"))

##Turning the dataset to long

df_long <- df_processed %>%
  pivot_longer(cols = starts_with("var3"), names_to = "date_column", values_to = "date") %>%
  mutate(Flash = ifelse(startsWith(date, "F"), "1", ""),
         date = ifelse(startsWith(date, "F"), gsub("^F ", "", date), date),
         date = gsub("\\;", "", date), 
         date = gsub("\\s+"," ",date)) %>%
  select(-date_column) %>%
  arrange(country)

df_long <- df_long %>% drop_na(date) %>% 
  filter(!grepl("\\?", date)) 

df_long <- df_long %>%
  filter(!grepl("benchmark", date)) %>%
  filter(!grepl("end", date)) %>%
  separate(date, into = c("day", "month", "release_number", "quarter", "year"), sep = " ") %>% 
  add_column(date = NA, 
             calendar_title = NA, 
             full_name = 'country', 
             calendar_description = NA, 
             in_calendar = NA, 
             id = NA)

df_long_edit = df_long %>% mutate(year = ifelse((quarter %in% c("Q3", "Q4")) & (month %in% c("Jan", "Feb", "Mar", "Apr", "May")),
                     as.integer(year) + 1,
                     year))

df_long_edit = select(df_long_edit, id, day, month, year, date, calendar_title, country, full_name, calendar_description, release_number, Flash, in_calendar)

df_long_edit$day <- ifelse(nchar(df_long_edit$day) < 2, paste0("0", df_long_edit$day), df_long_edit$day)
df_long_edit$year <- ifelse(nchar(df_long_edit$year) == 2, paste0("20", df_long_edit$year), df_long_edit$year)

month_numbers <- sprintf("%02d", seq_along(month.abb))
df_long_edit$month <- month_numbers[match(df_long_edit$month, month.abb)]


# Concatenate day, month, and year into the date column
df_long_edit$date <- paste(df_long_edit$year, df_long_edit$month, df_long_edit$day, sep = "")

# Copy the contents of the country column into the full_name column
df_long_edit$full_name <- df_long_edit$country

# Replace contents of country column with ISO code 3

df_long_edit$country <- countrycode(df_long_edit$country, origin = "country.name", destination = "iso3c")
df_long_edit$country <- ifelse(df_long_edit$full_name == "Euro area", "EMU", df_long_edit$country)
df_long_edit$country <- ifelse(df_long_edit$full_name == "European Union", "EUR", df_long_edit$country)


# Concatenate country, release_number, and flash into calendar_title column
df_long_edit$Flash <- ifelse(df_long_edit$Flash == 1, "f", df_long_edit$Flash)
df_long_edit$calendar_title <- paste(df_long_edit$country, df_long_edit$Flash, df_long_edit$release_number, sep = "")
df_long_edit$id <- paste(df_long_edit$date, df_long_edit$calendar_title, sep = "")

# Check if observation is duplicated in file "file.xlsx" and update in_calendar column
existing_data <- read_excel("C:/Users/sacksferrari_s/OneDrive - OECD/R Projects and Files/updated qna calendar_2.xlsx")
existing_data$year = as.character(existing_data$year)
df_long_edit$in_calendar <- ifelse(df_long_edit$id %in% existing_data$id, "Yes", "No")

df_long_edit_filter <- df_long_edit[df_long_edit$in_calendar == "No", ]
# Append filtered rows to "old_data"
appended_data <- rbind(existing_data, df_long_edit_filter)
openxlsx::write.xlsx(appended_data, "C:/Users/sacksferrari_s/OneDrive - OECD/R Projects and Files/updated qna calendar_2.xlsx", col_names=T)

###Now, we run this: https://ssacksferr.shinyapps.io/calendar_app/
