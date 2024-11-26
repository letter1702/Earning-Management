#install.packages("stargazer")
setwd("C:/Users/DELL/OneDrive/Máy tính/Tài liệu/tài liệu uel/Năm 3/HK2/Gói 2/ck/K214140957")
#Load the package
library(readxl)
library(dplyr)
library(openxlsx)
library(ggplot2)
library(corrplot)
library(plm)
library(stargazer)
#Đọc sheet
data <- read_excel("data (xuất từ R).xlsx")
#Thống kê mô tả
summary_table <- data.frame(
  Variable = c("EM", "ROA", "OCF", "LEV", "LOSS", "GA"),
  Min = c(min(data$EM), min(data$ROA), min(data$OCF), min(data$LEV), min(data$LOSS), min(data$GA)),
  "1st Qu." = c(quantile(data$EM, 0.25), quantile(data$ROA, 0.25), quantile(data$OCF, 0.25), quantile(data$LEV, 0.25), quantile(data$LOSS, 0.25), quantile(data$GA, 0.25)),
  Median = c(median(data$EM), median(data$ROA), median(data$OCF), median(data$LEV), median(data$LOSS), median(data$GA)),
  Mean = c(mean(data$EM), mean(data$ROA), mean(data$OCF), mean(data$LEV), mean(data$LOSS), mean(data$GA)),
  "3rd Qu." = c(quantile(data$EM, 0.75), quantile(data$ROA, 0.75), quantile(data$OCF, 0.75), quantile(data$LEV, 0.75), quantile(data$LOSS, 0.75), quantile(data$GA, 0.75)),
  Max = c(max(data$EM), max(data$ROA), max(data$OCF), max(data$LEV), max(data$LOSS), max(data$GA))
)
for (col in names(summary_table)) {
  if (is.numeric(summary_table[[col]])) {
    summary_table[[col]] <- sprintf("%.5f", summary_table[[col]])
  }
}
# Hiển thị bảng
summary_table

#Trực quan hóa dữ liệu
cor_matrix <- cor(data[, c("EM", "ROA", "OCF", "LEV", "LOSS", "GA")])
round(cor_matrix, 2)
#Create the heatmap
corrplot(cor_matrix, method = "color", type = "lower", 
         tl.col = "black", tl.srt = 45, 
         addCoef.col = "black", number.cex = 0.8,
         col = colorRampPalette(c("#BB4444", "#EE9988", "#FFFFFF", "#77AADD", "#4477AA"))(200))

stargazer(cor_matrix, type = "latex", title = "Thống kê mô tả các biến nghiên cứu") #output latex

# Tạo bộ data mới với các cột cần thiết
dt <- data[, c("GICS Industry Name","Code", "Date","asset","NI", "EM", "ROA", "OCF", "LEV", "LOSS", "GA")]
dt$year <- as.numeric(format(as.Date(dt$Date), "%Y"))
dt <- dt[, -which(names(dt) == "Date")]
dt <- dt %>%
  relocate(year, .after = Code)
names(dt)[names(dt) == "GICS Industry Name"] <- "Industry"
#Tính EM, asset, NI trung bình theo năm
result_table <- dt %>%
  group_by(year) %>%
  summarize(
    em_mean = mean(EM),
    assset_mean = mean(asset) / 1e12,
    NI_mean = mean(NI)/ 1e12
  )
#Tổng tài sản trung bình các công ty niêm yết giai đoạn 2015 - 2023
ggplot(result_table, aes(x = year, y = assset_mean)) +
  geom_bar(stat = "identity", fill = "steelblue") +
  labs(x = "Năm", y = "(Nghìn tỷ)") +
  scale_x_continuous(breaks = seq(min(result_table$year), max(result_table$year), by = 1)) +
  theme_minimal()+
  theme(
    plot.title = element_text(size = 13),
    axis.title = element_text(size = 13),
    axis.text = element_text(size = 12),
    axis.text.x = element_text(angle = 45, hjust = 1)
  )
#LNST trung bình các công ty niêm yết giai đoạn 2015 - 2023
ggplot(result_table, aes(x = year, y = NI_mean)) +
  geom_bar(stat = "identity", fill = "steelblue") +
  labs(x = "Năm", y = "(Nghìn tỷ)") +
  scale_x_continuous(breaks = seq(min(result_table$year), max(result_table$year), by = 1)) +
  theme_minimal()+
  theme(
    plot.title = element_text(size = 13),
    axis.title = element_text(size = 13),
    axis.text = element_text(size = 12),
    axis.text.x = element_text(angle = 45, hjust = 1)
  )

# Vẽ đồ thị EM
ggplot(result_table, aes(x = year, y = em_mean)) +
  geom_line(color = "steelblue", size = 1.5) +
  geom_point(color = "darkred", size = 3) +
  labs(
    x = "Năm",
    y = "EM"
  ) +
  scale_x_continuous(breaks = seq(min(result_table$year), max(result_table$year), by = 1)) +
  theme_minimal() +
  theme(
    plot.title = element_text(size = 16, face = "bold"),
    axis.title = element_text(size = 14),
    axis.text = element_text(size = 12),
    axis.text.x = element_text(angle = 45, hjust = 1)
  )

#Fixed Effects Model
fe1 <- plm(EM ~ ROA + OCF + LEV + LOSS + GA, data = dt, index = c("Code","year"),model = "within")
summary(fe1)
fixef(fe1)
#Ảnh hưởng covid_19
dt$COVID_19 <- ifelse(dt$year %in% c(2020, 2021), 1, 0)
# Giả sử bạn có data frame tên 'dt' với cột 'year'
dt$Cases <- 0
# Gán giá trị cho từng năm
dt$Cases[dt$year == 2020] <- 1737/96648685
dt$Cases[dt$year == 2021] <- 2746628/97468029 
#dt$Cases[dt$year == 2022] <- 11525231/98186856
#dt$Cases[dt$year == 2023] <- 11624114/98858950

#fe2 <- plm(EM ~ COVID_19 + factor(year), data = dt,index = c("Code"), model = "within")
#summary(fe2)
#fe3 <- plm(EM ~ Cases + factor(year), data = dt,index = c("Code"), model = "within")
#summary(fe3)

fem1 <- plm(EM ~ COVID_19 + ROA + OCF + LEV + LOSS + GA, data = dt, index = c("Code","year"),model = "within")
fem2 <- plm(EM ~ Cases + ROA + OCF + LEV + LOSS + GA, data = dt, index = c("Code","year"),model = "within")

#Xuất qua latex
stargazer(fe1, type = "text", ci.level = 0.95, title = "Kết quả mô hình fixed effects") #output text
stargazer(fe1, type = "latex", ci.level = 0.95, title = "Kết quả mô hình fixed effects") #output latex
stargazer(fem1,fem2, type = "text",ci.level = 0.95, title = "Đại dịch COVID-19 và hoạt động quản trị lợi nhuận")
stargazer(fem1,fem2, type = "latex",ci.level = 0.95, title = "Đại dịch COVID-19 và hoạt động quản trị lợi nhuận")
