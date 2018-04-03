
parse.bp = function() {
  # The function converts selected sheets
  # from the BP Statistical Review of World Energy 2017
  # provided as an Excel file here
  # https://www.bp.com/content/dam/bp/en/corporate/excel/energy-economics/statistical-review-2017/bp-statistical-review-of-world-energy-2017-underpinning-data.xlsx
  # into a format that can be easier used for data analysis
  # It is assumed that the Excel file is saved in your working directory
  # under the name "bp.xlsx"
  # You can run this function step by step

  library(readxl)
  library(openxlsx)
  library(tidyr)
  library(restorepoint)
  library(dplyr)
  
  file = "bp.xlsx"
  
  sheets = getSheetNames(file)
  
  #writeClipboard(paste0('"',sheets,'"', collapse=", "))
  
  used.sheets = c("Primary Energy Consumption", "Oil Production - Tonnes",  "Oil Consumption - Tonnes", "Oil - Refinery capacities", "Gas - Proved reserves history ", "Gas Production - Mtoe", "Gas Consumption - Mtoe", "Coal Production - Mtoe", "Coal Consumption -  Mtoe", "Nuclear Consumption - TWh", "Nuclear Consumption - Mtoe", "Hydro Consumption - TWh", "Hydro Consumption - Mtoe", "Other renewables -TWh", "Other renewables - Mtoe", "Solar Consumption - TWh", "Solar Consumption - Mtoe", "Wind Consumption - TWh ", "Wind Consumption - Mtoe", "Geo Biomass Other - TWh", "Geo Biomass Other - Mtoe",  "Biofuels Production - Ktoe", "Electricity Generation ", "Carbon Dioxide Emissions", "Geothermal capacity", "Solar capacity", "Wind capacity")
  
  sheet.nums = match(used.sheets, sheets)
  
  
  parse.bp.sheet = function(i=1) {
    restore.point("parse.bp.sheet")
    s = sheet.nums[i]
    sheet.name = used.sheets[i]
    xl = read_excel("bp.xlsx", sheet=s)
    
    na.rows = is.na(xl[[1]]) 
    
    xl = xl[!na.rows,]
    has.subtitle = is.na(xl[1,2])
    
    units = as.character(xl[1+has.subtitle,1])
    
    na.rows = rowSums(is.na(xl[,-1])) == (NCOL(xl)-1) 
    xl = xl[!na.rows,]
    
    colnames(xl) = c("region",xl[1,-1])
    xl = xl[-1,]
    dupl= duplicated(colnames(xl))
    xl = xl[!dupl]
    
    dat = gather(xl,key=year, value=value, -region)
    
    dat$year = as.integer(dat$year)
    dat = filter(dat, !is.na(dat$year))
    #dat$value = as.numeric(dat$value)
    dat$item = sheet.name
    dat$unit = units
    dat  
  }

  li = lapply(seq_along(used.sheets), parse.bp.sheet)
  df = bind_rows(li)
  unique(df$unit)
  regions = unique(df$region)
  
  replace.region = function(old, new) {
    rows = df$region == old
    df$region[rows] <<- new
  }
  
  writeClipboard(paste0('replace.region("', regions,'","',regions,'")', collapse="\n"))
  replace.region("France (Guadeloupe)","")
  replace.region("Portugal (The Azores)","")
  replace.region("USSR","")
  replace.region("Russia (Kamchatka)","")
  replace.region("Other S. & Cent. America","")
  replace.region("Other Europe & Eurasia","")
  replace.region("China Hong Kong SAR","China Hong Kong SAR")
  replace.region("Other Asia Pacific","")
  replace.region("Total World","Total World")
  replace.region("of which: OECD","OECD")
  replace.region("Non-OECD","Non-OECD")
  replace.region("European Union #","EU")
  replace.region("European Union","EU")
  replace.region("wLess than 0.05%.","")
  df = filter(df, region!="")
  
  library(countrycode)
  regions = unique(df$region)
  cntry = countrycode(regions, "country.name", "iso3c")
  is_cntry = !is.na(cntry)
  continent = countrycode(cntry, "iso3c","continent")
  wbregion = countrycode(cntry, "iso3c","region")
  
  cntry[!is_cntry] = regions[!is_cntry]
  
  mrows = match(df$region,regions)
  df$cntry = cntry[mrows]
  
  df$is_cntry = is_cntry[mrows]
  df$continent = continent[mrows]
  df$country = df$region
  df$region = wbregion[mrows]

   
  
  
  # Transform into wide format
  library(tidyr)
  library(stringtools)
  df$var = paste0(str.left.of(df$item,"-"),", ",df$unit)
  
  dw = df %>%
    select(year, cntry, country, continent, region, is_cntry, var, value) %>%
    tidyr::spread(key=var, value=value)

  # Investigate duplicated
  #dupl = duplicated(select(dw, country, year))
  #du = dw[dupl,c("country","year")]
  #du = inner_join(du, dw, by=c("country","year")) %>%
  #  arrange(country, year)
    
   
  # Add real gdp and population from the Penn World Tables
  library(pwt9)
  data("pwt9.0")
  pwt = pwt9.0
  
  pw = transmute(pwt, cntry=isocode,year=year, pop=pop, gdp=rgdpe / 1000, gdp_pc=gdp / pop*1000) 
  
  dw = right_join(pw,dw, by=c("cntry","year")) %>%
    filter(!is.na(country)) %>%
    select(cntry, region, year, pop, gdp, gdp_pc, everything())
  
  # Update missing gdp per capita and pop numbers
  impute = function(x) {
    x = ifelse(!is.na(x),x, lag(x) + (lag(x,1)-lag(x,2)) )
  }
  dw = group_by(dw, cntry)
  dw = mutate(dw, pop=impute(pop), gdp=impute(gdp))
  dw = mutate(dw, pop=impute(pop), gdp=impute(gdp))
  dw = mutate(dw, pop=impute(pop), gdp=impute(gdp))
  dw = ungroup(dw)
  dw = mutate(dw, gdp_pc=gdp / pop*1000)
   
  
  units = unique(df$unit)
  
  writeClipboard(paste0('"', units,'", ','"', units,'", 1e6'))

  txt = '"unit","pc_unit","factor",
"Million tonnes oil equivalent", "Tonnes oil equivalent", 1e6
"Million tonnes", "Tonnes", 1e6
"Thousand barrels daily*", "Barrels daily*", 1e3
"Trillion cubic metres", "Million cubic metres", 1e6
"Terawatt-hours", "Kilowatt-hours", 1e9
"Thousand tonnes of oil equivalent", "Kilogram of oil equivalent", 1e6
"Million tonnes carbon dioxide", "Tonnes carbon dioxide", 1e6
"Megawatts", "Watts", 1e6'
  to.pc = read.csv(textConnection(txt),stringsAsFactors = FALSE)
  
  cols = colnames(dw)[-(1:9)]
  for (col in cols) {
    dw[[col]] = as.numeric(dw[[col]])
  }
  
  pc = vector("list",NROW(cols))
  names(pc) = cols
  for (i in seq_along(cols)) {
    col = cols[i]
    val = dw[[col]]
    unit = str.trim(str.right.of(col,", "))
    row = match(unit, to.pc$unit)
    val = val * (to.pc$factor[row]/1e6) / dw$pop
    pc[[i]]  = val
    names(pc)[i] = paste0(str.trim(str.left.of(col,","))," per capita, ",to.pc$pc_unit[row])
  }
  pc = as_data_frame(pc)  

  dw = cbind(dw,pc)
  
  dw = arrange(dw, cntry, year)
  
  saveRDS(dw, "bp.Rds")
  readr::write_csv(dw, "bp.csv")
  
}


motion.plot.solar.wind = function() {
  # An example of creating a motin plot from the data

  dat = readr::read_csv("bp.csv")
  dat$iso3c = dat$cntry
  
  cols = colnames(dat)
  cols = cols[has.substr(cols, "capita")]
  writeClipboard(paste0('"', cols,'"', collapse=","))

  dat = filter(dat, is_cntry)
  
  cols = c("iso3c", "continent","country", "region","year","pop","gdp","gdp_pc","Carbon Dioxide Emissions per capita,  Tonnes carbon dioxide","Electricity Generation per capita,  Kilowatt-hours","Solar capacity per capita,  Watts","Solar Consumption per capita,  Kilowatt-hours","Wind capacity per capita,  Watts","Wind Consumption per capita,  Kilowatt-hours",
 "Geothermal capacity per capita,  Watts","Hydro Consumption per capita,  Kilowatt-hours","Nuclear Consumption per capita,  Kilowatt-hours","Other renewables per capita,  Kilowatt-hours"
    
    )
  
  df = dat[,cols]
  library(googleVis)
  
  p = gvisMotionChart(df, idvar = "iso3c",timevar = "year",xvar = "Wind capacity per capita,  Watts", yvar="Solar capacity per capita,  Watts",colorvar = "continent",sizevar = "pop")
  
  plot(p)
}