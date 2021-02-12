library(tidyverse)
library(wbstats)
library(officer)


# Create df using wbstats
# create new variable so that that gdp_pct_growth is formatted as %
gdp_pct_growth_df <- wbstats::wb_data(country = "countries_only", 
                                      indicator = "NY.GDP.PCAP.KD.ZG", 
                                      start_date = 2009,
                                      end_date = 2019, 
                                      return_wide = FALSE) %>% 
  mutate(gdp_pct_growth = value/100)

# Create a df with only data for France
df <- filter(gdp_pct_growth_df, country == "France")

# Create a simple line chart showing trends from 2009-2019
# geom_hline marks 0% growth for reading ease
# axis.title.x= element_blank() removes the x-axis (Year) title label
# axis.text.x = angle = 90 rotates the axis x text

ggplot(df, aes(x=date, y=gdp_pct_growth)) +
  geom_line(size = 1)+ 
  geom_point(size = 3) + 
  geom_hline(yintercept=0, linetype="dashed",  color = "red", size=1)  +
  scale_x_continuous("Year", breaks = seq(2009,2019, 1), limits = c(2009, 2019))+ 
  scale_y_continuous("GDP per capita growth (annual %)", labels = scales::percent)+
  labs(title = "GDP per capita growth (annual %)",
       subtitle = "France, 2009-2019",
       caption = "Source: World Bank Open Data")+
  theme_minimal()+ 
  theme(axis.title.x= element_blank(),
        axis.text.x = element_text(angle = 90))


# Make a list of all unique countries
country_list <- unique(gdp_pct_growth_df$country)

file_path_list <- list()

for( i in seq_along(country_list)) {
  
  # For each iteration, create a df for only that country
  
  df <- filter(gdp_pct_growth_df, country == country_list[i])
  
  # For each iteration, create a plot
  # Use paste0(counter_list[i]), to change labels when needed
  
  plot <- ggplot(df, aes(x=date, y=gdp_pct_growth)) +
    geom_line(size = 1)+ 
    geom_point(size = 3) + 
    geom_hline(yintercept=0, linetype="dashed",  color = "red", size=1)  +
    scale_x_continuous("Year", breaks = seq(2009,2019, 1), limits = c(2009, 2019))+ 
    scale_y_continuous("GDP per capita growth (annual %)", labels = scales::percent)+
    labs(title = "GDP per capita growth (annual %)",
         subtitle = paste0(country_list[i], ", ", "2009-2019"),
         caption = "Source: World Bank Open Data")+
    theme_minimal()+ 
    theme(axis.title.x= element_blank(),
          axis.text.x = element_text(angle = 90))
  
  # Store file path
  # str_replace_all replaces countries names so that it is alpha numeric to avoid invalid file path issues. 
  
  plot_file_path <- paste0(getwd(),"/graphs/", str_replace_all(country_list[i], "[^[:alnum:]]","_"), ".png")
  
  # Append filepath to list for each iteration
  file_path_list <- append(file_path_list, list(plot_file_path))
  
  # Save the graph
  ggsave(filename= plot_file_path, plot=plot, width = 7, height = 5, units = "in", scale = 1, dpi = 300)
  
}
                
# Put it in the pptx
# Read in the blank pptx
blank_ppt <- read_pptx("template.pptx")


# Loop to add slides 

for (u in seq(file_path_list)){
  gdp_growth_ppt <- add_slide(blank_ppt, 
                              layout = "Graph",
                              master = "Office Theme") %>%
    ph_with(value = external_img(file_path_list[[u]]),
            location = ph_location_type(type = "body"))}

gdp_growth_ppt %>% 
  print(target = "gdp_per_capita_%_growth_slide_deck.pptx") 


# Using lapply
# Create a function to make the graph 

country_graph <- function(x) {
  
  ### X argument is country name
  
  df <- filter(gdp_pct_growth_df, country == x)
  
plot <- ggplot(df, aes(x=date, y=gdp_pct_growth)) +
    geom_line(size = 1)+ 
    geom_point(size = 3) + 
    geom_hline(yintercept=0, linetype="dashed",  color = "red", size=1)  +
    scale_x_continuous("Year", breaks = seq(2009,2019, 1), limits = c(2009, 2019))+ 
    scale_y_continuous("GDP per capita growth (annual %)", labels = scales::percent)+
    labs(title = "GDP per capita growth (annual %)",
         subtitle = paste0(x, ", ", "2009-2019"),
         caption = "Source: World Bank Open Data")+
    theme_minimal()+ 
    theme(axis.title.x= element_blank(),
          axis.text.x = element_text(angle = 90))

### Store file path
### str_replace_all replaces countries names so that it is alpha numeric to avoid invalid file path issues. 

plot_file_path <- paste0(getwd(),"/graphs_function/", str_replace_all(x, "[^[:alnum:]]","_"), ".png")

### Append filepath to list for each iteration
ggsave(filename= plot_file_path, plot=plot, width = 7, height = 5, units = "in", scale = 1, dpi = 300)

plot_file_path

}

### Use map to make and save graphs than make a list of file paths 
country_graphs_list <- country_list %>% 
  purrr::set_names() %>% 
  purrr::map(country_graph)


### pptx loop
blank_ppt <- read_pptx("template.pptx")

for (u in seq(country_graphs_list)){
  gdp_growth_ppt <- add_slide(blank_ppt, 
                              layout = "Graph",
                              master = "Retrospect") %>%
    ph_with(value = external_img(country_graphs_list[[u]]),
            location = ph_location_type(type = "body"))}

gdp_growth_ppt %>% 
  print(target = "gdp_per_capita_%_growth_slide_deck.pptx") 




