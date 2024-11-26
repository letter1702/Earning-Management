setwd("C:/Users/DELL/OneDrive/Máy tính/Tài liệu/tài liệu uel/Năm 3/HK2/Gói 2/ck/K214140957")
#Load the package
library(readxl)
library(dplyr)
library(openxlsx)

#Đọc sheet
data_goi <- read_excel("K214140957.xlsx", sheet = "Sheet2")
comp <- read_excel("K214140957.xlsx", sheet = "Sheet1")

# Assuming the data frame is named 'data_comp'
data_comp <- comp[, c("Code", "Company Common Name", "GICS Industry Name")]
data <- full_join(data_comp,data_goi, by = "Code")
#Đổi tên cột
names(data)[names(data) == "Total Assets"] <- "asset"
names(data)[names(data) == "Revenue from Business Activities - Total"] <- "revenue"
names(data)[names(data) == "Loans & Receivables - Net - Short-Term"] <- "receivables"
names(data)[names(data) == "Net Cash Flow from Operating Activities"] <- "CF"
names(data)[names(data) == "Property Plant & Equipment - Gross - Total"] <- "PPE_t"
names(data)[names(data) == "Total Liabilities"] <- "liabilities"
names(data)[names(data) == "Income Available to Common Shares"] <- "NI"

#Check data types
str(data)

#Convert non-numeric columns to numeric
data$asset = as.numeric(gsub("[^0-9.-]", "", data$asset))
data$revenue = as.numeric(gsub("[^0-9.-]", "", data$revenue))
data$receivables = as.numeric(gsub("[^0-9.-]", "", data$receivables))
data$CF = as.numeric(gsub("[^0-9.-]", "", data$CF))
data$PPE_t = as.numeric(gsub("[^0-9.-]", "", data$PPE_t))
data$liabilities = as.numeric(gsub("[^0-9.-]", "", data$liabilities))
data$NI = as.numeric(gsub("[^0-9.-]", "", data$NI))

#Tính các biến
data$ROA <- data$NI / data$asset
data$OCF <- data$CF / data$asset
data$LEV <- data$liabilities / data$asset
data$LOSS <- ifelse(data$NI < 0, 1, 0)

data <- data %>%
  group_by(Code) %>%
  mutate(
    GA = (asset - lag(asset))/lag(asset),
    delta_rev = revenue - lag(revenue),
    delta_rec = receivables - lag(receivables),
    A_t1 = lag(asset),  # Tài sản kỳ trước
    TA_t = NI - CF,
    dependent_var = TA_t / A_t1,  # Biến phụ thuộc
    inv_A_t1 = 1 / A_t1,  # Biến độc lập 1
    delta_rev_rec = (delta_rev - delta_rec) / A_t1,  # Biến độc lập 2
    ppe = PPE_t / A_t1  # Biến độc lập 3
  ) %>%
  filter(!is.na(dependent_var) & !is.na(inv_A_t1) & !is.na(delta_rev_rec) & !is.na(ppe)) %>%  # Lọc bỏ các giá trị NA
  ungroup()

#Thực hiện hồi quy OLS
model <- lm(dependent_var ~ inv_A_t1 + delta_rev_rec + ppe, data = data)

#Xem kết quả hồi quy
summary(model)
summary(model)$coefficients
# Trích xuất các hệ số alpha từ kết quả hồi quy
alpha_1 <- coef(model)["inv_A_t1"]
alpha_2 <- coef(model)["delta_rev_rec"]
alpha_3 <- coef(model)["ppe"]

#Tính giá trị EM (Earning Management)
data$NDA <- alpha_1*data$inv_A_t1 + alpha_2*data$delta_rev_rec + alpha_3*data$ppe
data$EM <- data$dependent_var - data$NDA

#write.xlsx(data, "C:/Users/DELL/OneDrive/Máy tính/Tài liệu/tài liệu uel/Năm 3/HK2/Gói 2/ck/data (xuất từ R).xlsx", rowNames = FALSE)

