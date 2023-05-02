

library(purrr)
library(readxl)
library(xlsx)

# Load the data from excel multiple sheet

sheet_names <- excel_sheets("C:/Users/Pythagorasweb/Documents/Data Analysis/R/Logistics/RandomForest/bankAnalysis/BankAnalysis/data/Data_Dictionary.xlsx")

# print sheet names to check the names of the different sheets

sheet_names

# Read all the sheets into one list

alldata <-lapply(sheet_names, function(x) {
  
  as.data.frame(read_excel("C:/Users/Pythagorasweb/Documents/Data Analysis/R/Logistics/RandomForest/bankAnalysis/BankAnalysis/data/Data_Dictionary.xlsx", sheet = x))
  
})



# Rename the list element

names(alldata) <- sheet_names


#=====Print the view, head and tail of each sheet

View(alldata$`Application Volume`)

head(alldata$`Application Volume`)

tail(alldata$`Application Volume`)



#====Print the view, head and tail of each sheet

View(alldata$`Competitor Rates`)

head(alldata$`Competitor Rates`)

tail(alldata$`Competitor Rates`)



# Print the view, head and tail of each sheet

View(alldata$Moodys)

head(alldata$Moodys)

tail(alldata$Moodys)

# =======Data cleaning and Exploration

# check for missing values

sum(is.na(alldata$Membership))

sum(is.na(alldata$`Application Volume`))

sum(is.na(alldata$`Competitor Rates`))

sum(is.na(alldata$Moodys))

# === NB: Application Volume has 412060 null values, hence we need to clean them up
# Check the location of the missing NA's

which(is.na(alldata$`Application Volume`))
# Compute the total missing values in each column
colSums(is.na(alldata$`Application Volume`))
# Remove all the missing values using the na.omit function
alldata$`Application Volume` <- na.omit(alldata$`Application Volume`)


# check if all the missing values are removed
sum(is.na(alldata$`Application Volume`))

#===ANALYSIS ON THE MEMBERSHIP SHEET=====

membership <- alldata$Membership
head(membership)
View(membership)
# Convert Month and Year to Date format
membership$Date <- as.Date(paste(membership$Year, membership$Month, "01", sep = "-"), format = "%Y-%b-%d")
View(membership)
# Plot the total membership over time
#plot(membership$Date, membership$Total, type = "l", xlab = "Date", ylab = "Total Membership", main = "Membership over Time")


# Calculate monthly and yearly totals
library(dplyr)
membership_monthly <- membership %>% 
  group_by(Date) %>% 
  summarize(Total = sum(Total))

membership_yearly <- membership %>% 
  group_by(Year) %>% 
  summarize(Total = sum(Total))

# Plot monthly and yearly totals
plot(membership_yearly$Year, membership_yearly$Total, type = "b", xlab = "Year", ylab = "Total Membership", main = "Yearly Membership Totals")


# === Calculate the year-over-year growth rate of the total membership
# Load the tidyquant package
library(tidyquant)
library(dplyr)

# Define the percent_change function
percent_change <- function(x) {
  100 * c(NA, diff(x) / head(x, -1))
}

## Calculate the year-over-year growth rate
yearly_growth <- membership_yearly %>% 
  mutate(percent_change = percent_change(Total)) %>% 
  filter(!is.na(percent_change)) %>% 
  mutate(yearly_change = c(NA, diff(percent_change)))


# Plot the yearly growth rate
plot(tail(yearly_growth$Year, -1), tail(yearly_growth$yearly_change, -1), 
     type = "b", 
     xlab = "Year", ylab = "Year-over-Year Growth Rate (%)", 
     main = "Year-over-Year Growth Rate of Total Membership",
     xlim = c(2008, 2022))

#The graph shows that in 2008 and 2009, the growth was sudden, and moved to 200%
# However, Between 2010 and 2021, there was no significant growth
#In 2022, the growth increased to 50%



#------Competitor rate-----
library(tidyverse)
library(lubridate)
competitor_rate <- alldata$`Competitor Rates`
View(competitor_rate)

str(competitor_rate)
summary(competitor_rate)

# Convert the reportdate column to a date format using the ymd() function 
competitor_rate <- competitor_rate %>% 
  mutate(reportdate = ymd(`Report Date`))

rate_summary <- competitor_rate %>% 
  group_by(Term, `Product Group`) %>% 
  summarize(mean_rate = mean(Rate, na.rm = TRUE), 
            median_rate = median(Rate, na.rm = TRUE),
            sd_rate = sd(Rate, na.rm = TRUE))

# visualization
library(ggplot2)
library(ggpubr)
ggplot(rate_summary, aes(x = Term, y = mean_rate, fill = `Product Group`)) + 
  geom_bar(stat = "identity", position = "dodge") + 
  labs(x = "Term", y = "Mean Rate", fill = "Product Group") + 
  ggtitle("Mean Competitor Rates by Term and Product Group")



#====== Application Volume===
application_volume <- alldata$`Application Volume`
View(application_volume)
str(application_volume)

# lookout for outliers
# create a boxplot of the volume variable
boxplot(application_volume$VOLUME, main = "Boxplot of VOLUME")

# visualize the volume
ggplot(application_volume, aes(x = VOLUME)) +
  geom_histogram()

# Visualize the relationship between loan volume and product group using a box plot.
ggplot(application_volume, aes(x = PRODUCT_GROUP, y = VOLUME)) +
  geom_boxplot()

# Calculate the average loan volume by product group using the aggregate() function.
avg_volume <- aggregate(application_volume$VOLUME, by = list(application_volume$PRODUCT_GROUP), mean)
colnames(avg_volume) <- c("Product Group", "Average Volume")

avg_volume

# using scattered pot: find the relationship between volume  and year entered
# create a scatter plot of VOLUME vs. YEARENTERED
ggplot(data = application_volume, aes(x = YEARENTERED, y = VOLUME)) + 
  geom_point() + 
  labs(title = "Scatter plot of VOLUME vs. YEARENTERED",
       x = "YEARENTERED",
       y = "VOLUME")


#relate the TERM_CODE and VOLUME
ggplot(data = application_volume, aes(x = TERM_CODE, y = VOLUME)) +
  geom_point() +
  labs(title = "Scatter Plot of VOLUME vs TERM_CODE",
       x = "TERM_CODE",
       y = "VOLUME")
