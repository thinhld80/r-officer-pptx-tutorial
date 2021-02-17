library(tidyverse)
library(wbstats)
library(officer)


# Create df using wbstats
# create new variable so that that gdp_pct_growth is formatted as %
gdp_pct_growth_df <- wbstats::wb_data(
  country = "countries_only",
  indicator = "NY.GDP.PCAP.KD.ZG",
  start_date = 2009,
  end_date = 2019,
  return_wide = FALSE) %>%
  mutate(gdp_pct_growth = value / 100)


# Make a list of all unique countries
country_list <- unique(gdp_pct_growth_df$country)

file_path_list <- list()

# loop to create graphs
for (i in seq_along(country_list)) {

  # For each iteration, create a df for only that country

  df <- filter(gdp_pct_growth_df, country == country_list[i])

  # For each iteration, create a plot
  # Use paste0(counter_list[i]), to change labels when needed

  plot <- ggplot(df, aes(x = date, y = gdp_pct_growth)) +
    geom_line(size = 1) +
    geom_point(size = 3) +
    geom_hline(yintercept = 0, linetype = "dashed", color = "red", size = 1) +
    scale_x_continuous("Year", breaks = seq(2009, 2019, 1), limits = c(2009, 2019)) +
    scale_y_continuous("GDP per capita growth (annual %)", labels = scales::percent) +
    labs(
      title = "GDP per capita growth (annual %)",
      subtitle = paste0(country_list[i], ", ", "2009-2019"),
      caption = "Source: World Bank Open Data"
    ) +
    theme_minimal() +
    theme(
      axis.title.x = element_blank(),
      axis.text.x = element_text(angle = 90)
    )

  # Store file path
  # str_replace_all replaces countries names so that it is alpha numeric to avoid invalid file path issues.

  plot_file_path <- paste0(getwd(), "/graphs/", str_replace_all(country_list[i], "[^[:alnum:]]", "_"), ".png")

  # Append filepath to list for each iteration
  file_path_list <- append(file_path_list, list(plot_file_path))

  # Save the graph
  ggsave(filename = plot_file_path, plot = plot, width = 7, height = 5, units = "in", scale = 1, dpi = 300)
}

# Read in the blank pptx
template_pptx <- read_pptx("template.pptx")

# Copy blank pptx to gdp_growth_pptx (which will be our final pptx)
gdp_growth_pptx <- template_pptx

# Loop to add slides
for (i in seq(file_path_list)) {
  plot_png <- external_img(file_path_list[[i]])

  add_slide(gdp_growth_pptx,
    layout = "Graph",
    master = "Retrospect") %>%
    ph_with(
      value = plot_png,
      location = ph_location_label(ph_label = "Content Placeholder 2")
    )
}

# Export
print(gdp_growth_pptx, target = "gdp_per_capita_%_growth_slide_deck.pptx")


# Using map
# Create a function to make the graph

country_graph <- function(x) {

  # X argument is country name

  df <- filter(gdp_pct_growth_df, country == x)

  ggplot(df, aes(x = date, y = gdp_pct_growth)) +
    geom_line(size = 1) +
    geom_point(size = 3) +
    geom_hline(yintercept = 0, linetype = "dashed", color = "red", size = 1) +
    scale_x_continuous("Year", breaks = seq(2009, 2019, 1), limits = c(2009, 2019)) +
    scale_y_continuous("GDP per capita growth (annual %)", labels = scales::percent) +
    labs(
      title = "GDP per capita growth (annual %)",
      subtitle = paste0(x, ", ", "2009-2019"),
      caption = "Source: World Bank Open Data"
    ) +
    theme_minimal() +
    theme(
      axis.title.x = element_blank(),
      axis.text.x = element_text(angle = 90)
    )
}

# Use map to make graphs
country_graphs_list <- country_list %>%
  purrr::set_names() %>%
  purrr::map(country_graph)


# Copy blank pptx to gdp_growth_pptx (which will be our final pptx)
gdp_growth_pptx <- template_pptx

# Loop to add slides
for (i in seq(country_graphs_list)) {
  add_slide(gdp_growth_pptx,
    layout = "Graph",
    master = "Retrospect") %>%
    ph_with(
      value = country_graphs_list[[i]],
      location = ph_location_type(type = "body")
    )
}

print(gdp_growth_pptx, target = "gdp_per_capita_%_growth_slide_deck_map.pptx")
