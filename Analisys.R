library(openxlsx)
library(lubridate)
library(dplyr)
library(ggplot2)
library(scales)
library(tidyr)
library(grid)
library(RColorBrewer)

space <- function(x, ...) {format(x, ..., big.mark = " ", scientific = FALSE, trim = TRUE)}

###Data reading and processing
#Read revenue and profit data for plants, select only several columns
marg <- read.xlsx("./data/DBmarginality.xlsx", sheet = 1)
marg <- tbl_df(data.frame(productGroup1 = marg[,50], productGroup2 = marg[,51], scenario = marg[,5],
                          month = marg[,16], volume = marg[,18], priceRub = marg[,19], usdRate = marg[,46],
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
        png(paste0("figures/",productN,".png"), width=3308/2, height=2339/2)
        p <- ggplot(filter(marg, productGroup2 == product), aes(x = month, y = priceUsd)) + 
                geom_point(aes(size = volume, colour = market), alpha = .7) +
                geom_line(aes(x = month, y = priceUsd, colour = market), data = filter(meanPrice, 
                                        productGroup2 == product), size = 1) +
                scale_color_brewer(palette = "Dark2", name  ="Средневзвешенная цена") +
                scale_size(breaks = c(100,1000,4000,8000,12000), range = c(2,10), name  ="Объем по сделке") +
                ggtitle(paste0(product,"\nзаключенные сделки по реализации")) +
                theme_bw(base_size = 18) + 
                theme(legend.position="bottom") +
                xlab(NULL) +
                ylab("Цена за тонну, USD") 
                
        print(p)                    
        dev.off()      
}

productList <- c("Аммиак", "Карбамид", "AN", "MAP/DAP", "NPKS", "NS", "АЗФ", "CAN/CNS", "NP",
               "Прочие")
productListNames <- c("Аммиак", "Карбамид", "AN", "MAP-DAP", "NPKS", "NS", "АЗФ", "CAN-CNS", "NP",
                 "Прочие")
lapply(productList, function(product) {makefigures(product, productListNames[which(productList == product)])})      