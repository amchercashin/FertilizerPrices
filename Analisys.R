library(openxlsx)
library(lubridate)
library(dplyr)
library(ggplot2)
library(scales)
library(tidyr)
library(grid)
library(RColorBrewer)
library(gridExtra)

space <- function(x, ...) {format(x, ..., big.mark = " ", scientific = FALSE, trim = TRUE)}

###Data reading and processing
#Read revenue and profit data for plants, select only several columns
marg <- read.xlsx("./data/DBmarginality.xlsx", sheet = 1)
marg <- tbl_df(data.frame(productGroup1 = marg[,50], productGroup2 = marg[,51], scenario = marg[,5],
                          month = marg[,16], volume = marg[,18], priceRub = marg[,27], usdRate = marg[,46],
                          ico = marg[,45], market = marg[,9]))

#data processing
marg$scenario <- toupper(marg$scenario) #all scenario charachters to upper case
marg$market <- tolower(marg$market)
marg$month <- dmy("01.01.1900") + days(marg$month) - days(2) #date to POSIXct

marg <- filter(marg, ico != "ВГО")
marg <- filter(marg, scenario == "ФАКТ") #Take only fact 
marg$scenario[marg$scenario == "ФАКТ"] <- "Факт" #back to normal spelling
marg$priceRub[is.na(marg$priceRub)] <- 0 #all NA revenue to zeros
marg$volume[is.na(marg$volume)] <- 0 #all NA revenue to zeros
marg <- filter(marg, volume > 1 & priceRub > 0)

marg <- mutate(marg, priceUsd = priceRub / usdRate)
marg <- select(marg, -(priceRub:ico))
meanPrice <- group_by(marg, productGroup2, month, market) %>% summarise(revenue = sum(volume * priceUsd), priceUsd = revenue / sum (volume))

#Summarise
#marg <- group_by(marg, scenario, companygroup, product, month) %>% summarise(marginal.profit = sum(marginal.profit))
#To narrow tidy form
#marg <- gather(marg, indicator, value, -(productGroup1:month))

makefigures <- function(product, productN) {
        png(paste0("figures/",productN,".png"), width=3308, height=2339)
        p <- ggplot(filter(marg, productGroup2 == product), aes(x = month, y = priceUsd)) + 
                geom_point(aes(size = volume, colour = market), alpha = .7) +
                geom_line(aes(x = month, y = priceUsd, colour = market), data = filter(meanPrice, 
                                        productGroup2 == product), size = 2) +
                scale_color_brewer(palette = "Dark2", name  ="Средневзвешенная цена") +
                scale_size(range = c(4,22), name  ="Объем по сделке") +
                scale_y_continuous(labels = space) +
                ggtitle(paste0(product,"\nзаключенные сделки по реализации")) +
                theme_bw(base_size = 32) + 
                theme(legend.position="top", plot.margin = unit(c(1, 1, 1, 2), "lines"), 
                      axis.text.y = element_text(angle = 90), panel.grid.major.x = element_line(size = 4),
                      panel.grid.minor.x = element_line(size = 3), legend.key.size = unit(2, "cm")) +
                xlab(NULL) +
                ylab("Цена FCA за тонну, USD") +
                scale_x_datetime(labels = date_format("%Y"), breaks = date_breaks("year"), minor_breaks = date_breaks("month"))
                
        v <- ggplot(filter(marg, productGroup2 == product), aes(x = month, y = volume)) + 
                #geom_line(stat = "summary", fun.y="sum") +
                geom_bar(aes(fill = market), alpha = .6, position = "stack", stat = "summary", fun.y="sum") +
                scale_fill_brewer(palette = "Dark2") +
                theme_bw(base_size = 32) + 
                theme(legend.position="none", plot.margin = unit(c(1, 1, 1, 2), "lines"), 
                      axis.text.y = element_text(angle = 90), axis.text.y = element_text(size=14),
                      panel.grid.major.x = element_line(size = 4), panel.grid.minor.x = element_line(size = 3)) +
                scale_y_continuous(labels = space) +
                xlab(NULL) +
                ylab("Объем реализации, т.") +
                scale_x_datetime(labels = date_format("%Y"), breaks = date_breaks("year"), minor_breaks = date_breaks("month")) 
        
        grid.arrange(p, v, nrow = 2, heights = c(3, 1))                  
        dev.off()      
}

productList <- c("Аммиак", "Карбамид", "AN", "MAP/DAP", "NPKS", "NS", "АЗФ", "CAN/CNS", "NP",
               "Прочие")
productListNames <- c("Аммиак", "Карбамид", "AN", "MAP-DAP", "NPKS", "NS", "АЗФ", "CAN-CNS", "NP",
                 "Прочие")
lapply(productList, function(product) {makefigures(product, productListNames[which(productList == product)])})      