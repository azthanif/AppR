# Memuat library yang diperlukan
# Pastikan Anda telah menginstal argonR dan argonDash:
# install.packages(c("argonR", "argonDash"))

library(shiny)
library(argonR)
library(argonDash)
library(DT)
library(ggplot2)
library(plotly)
library(corrplot)
library(car)
library(nortest)
library(dplyr)
library(stringr)
library(readr)
library(shinycssloaders)
library(officer)
library(flextable)
library(webshot)
library(tidyr)
library(gridExtra)
library(openxlsx)
library(leaflet)
library(RColorBrewer)
library(viridis)
library(colourpicker)
library(rsconnect)
rsconnect::deployApp("C:/Penyimpanan Utama/Document/HANIF/STIS/II STIS/Kuliah Genap STIS/KOMPUTASI STATISTIK/PROJEK UAS/UASKOMSTAT/PROJEKUAS")

# Memuat data dari URL
tryCatch({
  # Data utama SoVI
  sovi_data <- read_csv("https://raw.githubusercontent.com/bmlmcmc/naspaclust/main/data/sovi_data.csv")
  
  # Matriks penimbang jarak
  distance_matrix <- read.csv("https://raw.githubusercontent.com/bmlmcmc/naspaclust/main/data/distance.csv", row.names = 1)
  
  print("✅ Data berhasil dimuat dari URL")
  print(paste("📊 SoVI Data:", nrow(sovi_data), "rows,", ncol(sovi_data), "columns"))
  print(paste("📏 Distance Matrix:", nrow(distance_matrix), "rows,", ncol(distance_matrix), "columns"))
  
}, error = function(e) {
  print("⚠️ Gagal memuat data dari URL, menggunakan data sampel")
  
  set.seed(123)
  sovi_data <- data.frame(
    ID = 1:100,
    County = paste("County", 1:100),
    State = sample(c("CA", "TX", "FL", "NY", "PA"), 100, replace = TRUE),
    Population = rnorm(100, 50000, 15000),
    Income = rnorm(100, 45000, 12000),
    Education = rnorm(100, 85, 10),
    Healthcare = rnorm(100, 75, 15),
    Housing = rnorm(100, 65, 20),
    Employment = rnorm(100, 80, 12),
    Infrastructure = rnorm(100, 70, 18),
    SoVI_Score = rnorm(100, 0, 1),
    Latitude = runif(100, 25, 49),
    Longitude = runif(100, -125, -65),
    Region = sample(c("North", "South", "East", "West"), 100, replace = TRUE),
    Urban_Rural = sample(c("Urban", "Rural"), 100, replace = TRUE)
  )
  
  n <- nrow(sovi_data)
  distance_matrix <- matrix(runif(n*n, 0, 100), nrow = n, ncol = n)
  diag(distance_matrix) <- 0
  distance_matrix <- as.data.frame(distance_matrix)
  names(distance_matrix) <- paste0("County_", 1:n)
})

# Pra-pemrosesan data
if(ncol(sovi_data) > 0) {
  char_cols <- sapply(sovi_data, is.character)
  if(any(char_cols)) {
    sovi_data[char_cols] <- lapply(sovi_data[char_cols], as.factor)
  }
  
  numeric_vars <- names(sovi_data)[sapply(sovi_data, is.numeric)]
  if(length(numeric_vars) > 2) {
    var1 <- numeric_vars[1]
    var2 <- numeric_vars[2]
    
    sovi_data[[paste0(var1, "_Cat")]] <- cut(sovi_data[[var1]], 
                                             breaks = 3, 
                                             labels = c("Low", "Medium", "High"))
    sovi_data[[paste0(var2, "_Cat")]] <- cut(sovi_data[[var2]], 
                                             breaks = 2, 
                                             labels = c("Group_A", "Group_B"))
  }
}

print("📋 Struktur data:")
print(str(sovi_data))


# UI dengan ArgonDash
ui <- argonDashPage(
  title = "SAVI Dashboard",
  author = "UAS Komputasi Statistik",
  description = "Sistem Analisis Visualisasi Interaktif",
  sidebar = argonDashSidebar(
    vertical = TRUE,
    skin = "light",
    background = "white",
    size = "md",
    id = "main_sidebar",
    brand_logo = "LOGO_DASH.png",
    brand_url = "https://www.stis.ac.id",
    argonSidebarMenu(
      argonSidebarItem(
        tabName = "home",
        icon = argonIcon("tv-2"),
        "Beranda"
      ),
      argonSidebarItem(
        tabName = "data_mgmt",
        icon = argonIcon("folder-17"),
        "Manajemen Data"
      ),
      argonSidebarItem(
        name = "Eksplorasi Data",
        icon = argonIcon("chart-bar-32"),
        argonSidebarItem(tabName = "descriptive", "Statistik Deskriptif"),
        argonSidebarItem(tabName = "visualization", "Visualisasi & Peta"),
        argonSidebarItem(tabName = "distance_analysis", "Analisis Matriks Jarak")
      ),
      argonSidebarItem(
        tabName = "assumptions",
        icon = argonIcon("check-bold"),
        "Uji Asumsi"
      ),
      argonSidebarItem(
        name = "Statistik Inferensia",
        icon = argonIcon("atom"),
        argonSidebarItem(tabName = "mean_tests", "Uji Rata-rata"),
        argonSidebarItem(tabName = "prop_tests", "Uji Proporsi"),
        argonSidebarItem(tabName = "var_tests", "Uji Varians")
      ),
      argonSidebarItem(
        tabName = "anova",
        icon = argonIcon("chart-pie-35"),
        "ANOVA"
      ),
      argonSidebarItem(
        tabName = "regression",
        icon = argonIcon("chart-line"),
        "Regresi Linear Berganda"
      )
    )
  ),
  header = argonDashHeader(
    gradient = TRUE,
    color = "primary",
    separator = TRUE,
    separator_color = "secondary",
    argonH1("Dashboard Analisis Indeks Kerentanan Sosial", display = 4, class = "text-white")
  ),
  body = argonDashBody(
    argonTabItems(
      # Halaman Beranda
      argonTabItem(
        tabName = "home",
        argonRow(
          argonColumn(
            width = 12,
            argonCard(
              width = 12,
              title = "🎓 Dashboard SAVI - Sistem Analisis Visualisasi Interaktif",
              status = "primary",
              shadow = TRUE,
              hover_shadow = TRUE,
              p("Dashboard ini dibuat untuk memenuhi persyaratan UAS Komputasi Statistik STIS - Semester Genap TA. 2024/2025"),
              br(),
              h4("📋 Metadata Dataset:"),
              tags$a(href = "https://raw.githubusercontent.com/bmlmcmc/naspaclust/main/data/sovi_data.csv", target = "_blank", "📊 Dataset SoVI (Indeks Kerentanan Sosial)"), br(),
              tags$a(href = "https://raw.githubusercontent.com/bmlmcmc/naspaclust/main/data/distance.csv", target = "_blank", "📏 Matriks Jarak (Distance Matrix)"), br(),
              tags$a(href = "https://www.sciencedirect.com/science/article/pii/S2352340921010180", target = "_blank", "📖 Publikasi ScienceDirect - Artikel Data"),
              br(),br(),
              h4("🎯 Tujuan Dashboard:"),
              tags$ul(
                tags$li("Menganalisis data Indeks Kerentanan Sosial (SoVI) secara komprehensif"),
                tags$li("Menerapkan berbagai metode statistik inferensia dan deskriptif"),
                tags$li("Melakukan visualisasi data interaktif"),
                tags$li("Mengimplementasikan uji asumsi statistik yang diperlukan"),
                tags$li("Menyediakan platform analisis yang user-friendly untuk pembelajaran")
              ),
              br(),
              downloadButton("download_intro_card", "Unduh Info Pengantar (Word)", class="btn-primary")
            )
          )
        ),
        argonRow(
          argonInfoCard(
            value = paste(nrow(sovi_data), "Observasi"),
            title = "DATASET SOVI",
            stat = paste(ncol(sovi_data), "Variabel"),
            description = paste(sum(sapply(sovi_data, is.numeric)), "Numerik,", sum(!sapply(sovi_data, is.numeric)), "Kategorikal"),
            icon = argonIcon("sound-wave"),
            icon_background = "danger",
            hover_lift = TRUE
          ),
          argonInfoCard(
            value = paste(nrow(distance_matrix), "x", ncol(distance_matrix)),
            title = "MATRIKS JARAK",
            stat = "Simetris",
            description = "Penimbang Jarak Spasial",
            icon = argonIcon("map-big"),
            icon_background = "warning",
            hover_lift = TRUE
          ),
          argonInfoCard(
            value = "10+ Fitur",
            title = "FITUR DASHBOARD",
            stat = "7 Tab Analisis",
            description = "Manajemen Data, Eksplorasi, Uji Asumsi, Inferensia, ANOVA, Regresi",
            icon = argonIcon("app"),
            icon_background = "success",
            hover_lift = TRUE
          )
        ),
        argonRow(
          argonColumn(width = 6,
                      argonCard(
                        width = 12,
                        title = "📊 Ringkasan Data SoVI",
                        status = "info",
                        shadow = TRUE,
                        argonTabSet(
                          id = "home_sovi_tabs",
                          card_wrapper = TRUE,
                          horizontal = TRUE,
                          circle = FALSE,
                          size = "sm",
                          width = 12,
                          iconList = list(
                            argonIcon("align-left-2"), 
                            argonIcon("chart-bar-32"),
                            argonIcon("books")
                          ),
                          argonTab(
                            tabName = "Statistik Dasar",
                            active = TRUE,
                            withSpinner(DT::dataTableOutput("home_sovi_summary")),
                            downloadButton("download_sovi_summary", "📄 Unduh Ringkasan (Word)")
                          ),
                          argonTab(
                            tabName = "Visualisasi Cepat",
                            withSpinner(plotOutput("home_quick_viz", height = "300px")),
                            downloadButton("download_quick_viz", "🖼️ Unduh Plot (JPG)")
                          ),
                          argonTab(
                            tabName = "Tautan Berguna",
                            h5("📚 Referensi dan Dokumentasi:"),
                            tags$ul(
                              tags$li(tags$a(href = "https://rinterface.github.io/argonDash/", target = "_blank", "📖 Dokumentasi ArgonDash")),
                              tags$li(tags$a(href = "https://ggplot2.tidyverse.org/", target = "_blank", "📊 Dokumentasi ggplot2"))
                            )
                          )
                        )
                      )
          ),
          argonColumn(width = 6,
                      argonCard(
                        width = 12,
                        title = "📏 Informasi Matriks Jarak",
                        status = "warning",
                        shadow = TRUE,
                        argonTabSet(
                          id = "home_dist_tabs",
                          card_wrapper = TRUE,
                          horizontal = TRUE,
                          circle = FALSE,
                          size = "sm",
                          width = 12,
                          iconList = list(
                            argonIcon("bullet-list-67"), 
                            argonIcon("chart-pie-35"),
                            argonIcon("settings-gear-65")
                          ),
                          argonTab(
                            tabName = "Info Dasar",
                            active = TRUE,
                            withSpinner(verbatimTextOutput("distance_info")),
                            downloadButton("download_distance_info", "📄 Unduh Info (Word)")
                          ),
                          argonTab(
                            tabName = "Statistik Jarak",
                            withSpinner(DT::dataTableOutput("distance_stats_table")),
                            downloadButton("download_distance_stats", "📊 Unduh Statistik (Excel)")
                          ),
                          argonTab(
                            tabName = "Penggunaan",
                            h5("🔍 Aplikasi Matriks Jarak:"),
                            p("Analisis Clustering, Analisis Spasial, Multidimensional Scaling (MDS), Deteksi Outlier, dll."),
                            br(),
                            downloadButton("download_usage_info", "Unduh Info Penggunaan (Word)", class="btn-warning")
                          )
                        )
                      )
          )
        ),
        argonRow(
          argonCard(
            width = 12,
            title = "💾 Unduhan Lengkap",
            status = "success",
            p("Unduh semua file yang diperlukan untuk analisis offline dan dokumentasi lengkap penelitian Anda."),
            argonRow(
              argonColumn(width=3, downloadButton("download_metadata_complete", "📖 Metadata Lengkap", class="btn-info")),
              argonColumn(width=3, downloadButton("download_sovi_data", "📊 Data SoVI", class="btn-success")),
              argonColumn(width=3, downloadButton("download_distance_matrix", "📏 Matriks Jarak", class="btn-primary")),
              argonColumn(width=3, downloadButton("download_combined_data", "💾 Data Gabungan", class="btn-warning"))
            )
          )
        )
      ),
      # Halaman Manajemen Data
      argonTabItem(
        tabName = "data_mgmt",
        argonRow(
          argonColumn(width = 4,
                      argonCard(
                        width = 12,
                        status = "primary",
                        title = "🔧 Pengaturan Kategorisasi Data",
                        selectInput("cat_variable", "Pilih Variabel untuk Dikategorikan:", choices = NULL),
                        numericInput("cat_groups", "Jumlah Kategori:", value = 3, min = 2, max = 5),
                        radioButtons("cat_method", "Metode Kategorisasi:",
                                     choices = list("Berdasarkan Kuantil" = "quantile", "Berdasarkan Range Sama" = "equal", "Berdasarkan K-means" = "kmeans")),
                        textInput("cat_labels", "Label Kategori (pisahkan dengan koma):", placeholder = "Rendah, Sedang, Tinggi"),
                        br(),
                        actionButton("apply_categorization", "🔄 Terapkan Kategorisasi", class = "btn-primary"),
                        br(), br(),
                        h4("💾 Download Hasil:"),
                        downloadButton("download_categorized_data", "📊 Data (Excel)", class="btn-success"),
                        downloadButton("download_cat_interpretation", "📝 Interpretasi (Word)", class="btn-info"),
                        downloadButton("download_cat_report", "📋 Laporan (Word)", class="btn-warning")
                      )
          ),
          argonColumn(width = 8,
                      argonCard(
                        width = 12,
                        status = "info",
                        title = "📊 Hasil Kategorisasi Data",
                        argonTabSet(
                          id = "datamgmt_tabs",
                          width = 12,
                          argonTab(
                            tabName = "Data Terkategorisasi", 
                            withSpinner(DT::dataTableOutput("categorized_data"))
                          ),
                          argonTab(
                            tabName = "Visualisasi Perbandingan", 
                            withSpinner(plotOutput("categorization_plot")),
                            br(),
                            downloadButton("download_cat_plot_jpg", "🖼️ Unduh Plot (JPG)", class="btn-default")
                          ),
                          argonTab(
                            tabName = "Statistik Kategori", 
                            withSpinner(DT::dataTableOutput("category_stats")),
                            br(),
                            downloadButton("download_cat_stats", "📊 Unduh Statistik (Excel)", class="btn-default")
                          ),
                          argonTab(
                            tabName = "Interpretasi Lengkap", 
                            withSpinner(verbatimTextOutput("categorization_interpretation"))
                          )
                        )
                      )
          )
        )
      ),
      # Halaman Statistik Deskriptif
      argonTabItem(
        tabName = "descriptive",
        argonRow(
          argonColumn(width = 4,
                      argonCard(
                        width = 12,
                        status = "primary",
                        title = "📊 Pengaturan Analisis Deskriptif",
                        checkboxGroupInput("desc_vars", "Pilih Variabel:", choices = NULL),
                        radioButtons("desc_type", "Jenis Analisis:", choices = list("Statistik Dasar" = "basic", "Statistik Lengkap" = "complete", "Analisis Distribusi" = "distribution")),
                        checkboxInput("desc_by_group", "Analisis per Grup", FALSE),
                        conditionalPanel(
                          condition = "input.desc_by_group",
                          selectInput("desc_group_var", "Variabel Grouping:", choices = NULL)
                        ),
                        actionButton("calc_descriptive", "📊 Hitung Statistik", class = "btn-primary"),
                        br(), br(),
                        h4("💾 Download Hasil:"),
                        downloadButton("download_desc_excel", "📊 Statistik (Excel)", class="btn-success"),
                        downloadButton("download_desc_interpretation", "📝 Interpretasi (Word)", class="btn-info"),
                        downloadButton("download_desc_report", "📋 Laporan (Word)", class="btn-warning")
                      )
          ),
          argonColumn(width = 8,
                      argonCard(
                        width = 12,
                        status = "info",
                        title = "📊 Hasil Analisis Deskriptif",
                        argonTabSet(
                          id = "desc_tabs",
                          width = 12,
                          argonTab(tabName = "Tabel Statistik", withSpinner(DT::dataTableOutput("descriptive_stats"))),
                          argonTab(
                            tabName = "Visualisasi Distribusi", 
                            withSpinner(plotOutput("descriptive_plots", height = "600px")),
                            br(),
                            downloadButton("download_desc_plots_jpg", "🖼️ Unduh Plot (JPG)", class="btn-default")
                          ),
                          argonTab(
                            tabName = "Perbandingan Grup", 
                            withSpinner(plotOutput("group_comparison_plot")),
                            br(),
                            downloadButton("download_group_comp_plot_jpg", "🖼️ Unduh Plot (JPG)", class="btn-default")
                          ),
                          argonTab(tabName = "Interpretasi", withSpinner(verbatimTextOutput("descriptive_interpretation")))
                        )
                      )
          )
        )
      ),
      # Halaman Visualisasi
      argonTabItem(
        tabName = "visualization",
        argonRow(
          argonColumn(width = 4,
                      argonCard(
                        width = 12,
                        status = "primary",
                        title = "🎨 Pengaturan Visualisasi",
                        selectInput("plot_x", "Variabel X:", choices = NULL),
                        selectInput("plot_y", "Variabel Y:", choices = NULL),
                        selectInput("plot_type", "Jenis Plot:",
                                    choices = list("Scatter Plot" = "scatter", "Line Plot" = "line", "Box Plot" = "box", "Violin Plot" = "violin", "Histogram" = "hist", "Density Plot" = "density", "Heatmap" = "heatmap")),
                        conditionalPanel(
                          condition = "input.plot_type == 'box' || input.plot_type == 'violin'",
                          selectInput("plot_group", "Variabel Grouping:", choices = NULL)
                        ),
                        colourInput("plot_color", "Warna Utama:", "#5e72e4"),
                        checkboxInput("add_trend", "Tambahkan Garis Trend", FALSE),
                        textInput("plot_title", "Judul Plot:", placeholder = "Masukkan judul custom"),
                        actionButton("generate_plot", "🎨 Buat Plot", class = "btn-success")
                      )
          ),
          argonColumn(width = 8,
                      argonCard(
                        width = 12,
                        status = "info",
                        title = "📊 Visualisasi Data",
                        argonTabSet(
                          id = "viz_tabs",
                          width = 12,
                          argonTab(tabName = "Plot Interaktif", withSpinner(plotlyOutput("interactive_plot", height = "500px"))),
                          argonTab(tabName = "Plot Statis", withSpinner(plotOutput("static_plot", height = "500px")))
                        )
                      )
          )
        )
      ),
      # Halaman Analisis Matriks Jarak
      argonTabItem(
        tabName = "distance_analysis",
        argonRow(
          argonColumn(width = 4,
                      argonCard(
                        width = 12,
                        status = "primary",
                        title = "📏 Pengaturan Analisis Matriks Jarak",
                        radioButtons("distance_analysis_type", "Jenis Analisis:",
                                     choices = list("Statistik Deskriptif Matriks" = "descriptive", "Clustering Berdasarkan Jarak" = "clustering", "Visualisasi Heatmap" = "heatmap", "Analisis Komponen Utama" = "pca")),
                        conditionalPanel(
                          condition = "input.distance_analysis_type == 'clustering'",
                          numericInput("n_clusters", "Jumlah Cluster:", value = 3, min = 2, max = 10),
                          selectInput("cluster_method", "Metode Clustering:", choices = list("K-means" = "kmeans", "Hierarchical" = "hierarchical", "PAM" = "pam"))
                        ),
                        actionButton("run_distance_analysis", "📏 Jalankan Analisis", class = "btn-primary"),
                        br(), br(),
                        h4("💾 Download Hasil:"),
                        downloadButton("download_distance_results", "📊 Hasil (Excel)", class="btn-success"),
                        downloadButton("download_distance_plot", "🖼️ Visualisasi (JPG)", class="btn-warning"),
                        downloadButton("download_distance_interpretation", "📝 Interpretasi (Word)", class="btn-info")
                      )
          ),
          argonColumn(width = 8,
                      argonCard(
                        width = 12,
                        status = "info",
                        title = "📏 Hasil Analisis Matriks Jarak",
                        argonTabSet(
                          id = "dist_tabs",
                          width = 12,
                          argonTab(tabName = "Hasil Analisis", withSpinner(verbatimTextOutput("distance_analysis_results")), withSpinner(DT::dataTableOutput("distance_results_table"))),
                          argonTab(tabName = "Visualisasi", withSpinner(plotOutput("distance_plot", height = "600px"))),
                          argonTab(tabName = "Heatmap Jarak", withSpinner(plotOutput("distance_heatmap", height = "600px"))),
                          argonTab(tabName = "Interpretasi", withSpinner(verbatimTextOutput("distance_interpretation")))
                        )
                      )
          )
        )
      ),
      # Halaman Uji Asumsi
      argonTabItem(
        tabName = "assumptions",
        argonRow(
          argonColumn(width = 6,
                      argonCard(
                        width = 12,
                        status = "primary",
                        title = "🔍 Uji Normalitas",
                        selectInput("norm_var", "Pilih Variabel:", choices = NULL),
                        checkboxGroupInput("norm_tests", "Pilih Uji:",
                                           choices = list("Shapiro-Wilk" = "shapiro", "Kolmogorov-Smirnov" = "ks", "Lilliefors" = "lillie", "Anderson-Darling" = "ad", "Jarque-Bera" = "jb"),
                                           selected = c("shapiro", "lillie")),
                        sliderInput("norm_alpha", "Tingkat Signifikansi:", min = 0.01, max = 0.1, value = 0.05, step = 0.01),
                        actionButton("run_normality", "🔍 Jalankan Uji Normalitas", class = "btn-primary"),
                        br(), br(),
                        h4("💾 Download Hasil:"),
                        downloadButton("download_norm_results", "📊 Hasil (Word)", class="btn-success"),
                        downloadButton("download_norm_plots", "🖼️ Plot (JPG)", class="btn-warning"),
                        downloadButton("download_norm_interpretation", "📝 Interpretasi (Word)", class="btn-info")
                      )
          ),
          argonColumn(width = 6,
                      argonCard(
                        width = 12,
                        status = "warning",
                        title = "⚖️ Uji Homogenitas",
                        selectInput("homo_var", "Variabel Numerik:", choices = NULL),
                        selectInput("homo_group", "Variabel Grouping:", choices = NULL),
                        radioButtons("homo_test", "Jenis Uji:", choices = list("Levene's Test" = "levene", "Bartlett's Test" = "bartlett", "Fligner-Killeen Test" = "fligner")),
                        sliderInput("homo_alpha", "Tingkat Signifikansi:", min = 0.01, max = 0.1, value = 0.05, step = 0.01),
                        actionButton("run_homogeneity", "⚖️ Jalankan Uji Homogenitas", class = "btn-warning"),
                        br(), br(),
                        h4("💾 Download Hasil:"),
                        downloadButton("download_homo_results", "📊 Hasil (Word)", class="btn-success"),
                        downloadButton("download_homo_interpretation", "📝 Interpretasi (Word)", class="btn-info")
                      )
          )
        ),
        argonRow(
          argonColumn(width = 6,
                      argonCard(width=12, title = "🔍 Hasil Uji Normalitas", status="info",
                                withSpinner(verbatimTextOutput("normality_results")),
                                br(),
                                withSpinner(plotOutput("normality_plots"))
                      )
          ),
          argonColumn(width = 6,
                      argonCard(width=12, title = "⚖️ Hasil Uji Homogenitas", status="info",
                                withSpinner(verbatimTextOutput("homogeneity_results")),
                                br(),
                                withSpinner(plotOutput("homogeneity_plots")),
                                br(),
                                downloadButton("download_homo_plot_jpg", "🖼️ Unduh Plot (JPG)", class="btn-default")
                      )
          )
        ),
        argonRow(
          argonCard(width=12, title = "📝 Interpretasi Gabungan Uji Asumsi", status="success",
                    withSpinner(verbatimTextOutput("assumptions_interpretation")),
                    br(),
                    downloadButton("download_assumptions_report", "📋 Laporan Uji Asumsi Lengkap (Word)", class="btn-success")
          )
        )
      ),
      # Halaman Uji Rata-rata
      argonTabItem(
        tabName = "mean_tests",
        argonRow(
          argonColumn(width = 4,
                      argonCard(
                        width = 12,
                        status = "primary",
                        title = "🧪 Pengaturan Uji Rata-rata",
                        selectInput("mean_test_type", "Jenis Uji:",
                                    choices = list("One Sample t-test" = "t_one", "Two Sample t-test" = "t_two", "Paired t-test" = "t_paired", "Welch t-test" = "welch")),
                        selectInput("mean_test_var", "Variabel:", choices = NULL),
                        conditionalPanel("input.mean_test_type == 't_one'", numericInput("mu_value", "Nilai μ₀:", value = 0)),
                        conditionalPanel("input.mean_test_type == 't_two' || input.mean_test_type == 'welch'", selectInput("mean_group_var", "Variabel Grouping:", choices = NULL)),
                        conditionalPanel("input.mean_test_type == 't_paired'", selectInput("paired_var", "Variabel Pasangan:", choices = NULL)),
                        radioButtons("mean_alternative", "Hipotesis Alternatif:", choices = list("Two-sided" = "two.sided", "Greater than" = "greater", "Less than" = "less")),
                        sliderInput("mean_alpha", "Tingkat Signifikansi:", min = 0.01, max = 0.1, value = 0.05, step = 0.01),
                        actionButton("run_mean_test", "🧪 Jalankan Uji", class = "btn-primary"),
                        br(), br(),
                        h4("💾 Download Hasil:"),
                        downloadButton("download_mean_results", "📊 Hasil (Word)", class="btn-success"),
                        downloadButton("download_mean_interpretation", "📝 Interpretasi (Word)", class="btn-info")
                      )
          ),
          argonColumn(width = 8,
                      argonCard(
                        width = 12,
                        status = "info",
                        title = "🧪 Hasil Uji Rata-rata",
                        argonTabSet(
                          id = "mean_tabs",
                          width = 12,
                          argonTab(tabName = "Hasil Uji", withSpinner(verbatimTextOutput("mean_test_results"))),
                          argonTab(
                            tabName = "Visualisasi", 
                            withSpinner(plotOutput("mean_test_plot", height = "400px")),
                            br(),
                            downloadButton("download_mean_plot_jpg", "🖼️ Unduh Plot (JPG)", class="btn-default")
                          ),
                          argonTab(tabName = "Interpretasi Lengkap", withSpinner(verbatimTextOutput("mean_test_interpretation")))
                        )
                      )
          )
        )
      ),
      # Halaman Uji Proporsi
      argonTabItem(
        tabName = "prop_tests",
        argonRow(
          argonColumn(width = 4,
                      argonCard(
                        width = 12,
                        status = "primary",
                        title = "📊 Pengaturan Uji Proporsi",
                        selectInput("prop_test_type", "Jenis Uji:",
                                    choices = list("One Sample Proportion Test" = "prop_one", "Two Sample Proportion Test" = "prop_two", "Chi-square Goodness of Fit" = "chisq_gof", "Chi-square Independence" = "chisq_indep")),
                        selectInput("prop_var", "Variabel Kategorik:", choices = NULL),
                        conditionalPanel("input.prop_test_type == 'prop_one'", numericInput("prop_value", "Proporsi Hipotesis:", value = 0.5, min = 0, max = 1, step = 0.01)),
                        conditionalPanel("input.prop_test_type == 'prop_two' || input.prop_test_type == 'chisq_indep'", selectInput("prop_group_var", "Variabel Grouping:", choices = NULL)),
                        sliderInput("prop_alpha", "Tingkat Signifikansi:", min = 0.01, max = 0.1, value = 0.05, step = 0.01),
                        actionButton("run_prop_test", "📊 Jalankan Uji", class = "btn-primary"),
                        br(), br(),
                        h4("💾 Download Hasil:"),
                        downloadButton("download_prop_results", "📊 Hasil (Word)", class="btn-success"),
                        downloadButton("download_prop_interpretation", "📝 Interpretasi (Word)", class="btn-info")
                      )
          ),
          argonColumn(width = 8,
                      argonCard(
                        width = 12,
                        status = "info",
                        title = "📊 Hasil Uji Proporsi",
                        argonTabSet(
                          id = "prop_tabs",
                          width = 12,
                          argonTab(tabName = "Hasil Uji", withSpinner(verbatimTextOutput("prop_test_results"))),
                          argonTab(tabName = "Tabel Kontingensi", withSpinner(DT::dataTableOutput("contingency_table"))),
                          argonTab(
                            tabName = "Visualisasi", 
                            withSpinner(plotOutput("prop_test_plot", height = "400px")),
                            br(),
                            downloadButton("download_prop_plot_jpg", "🖼️ Unduh Plot (JPG)", class="btn-default")
                          ),
                          argonTab(tabName = "Interpretasi", withSpinner(verbatimTextOutput("prop_test_interpretation")))
                        )
                      )
          )
        )
      ),
      # Halaman Uji Varians
      argonTabItem(
        tabName = "var_tests",
        argonRow(
          argonColumn(width = 4,
                      argonCard(
                        width = 12,
                        status = "primary",
                        title = "📈 Pengaturan Uji Varians",
                        selectInput("var_test_type", "Jenis Uji:",
                                    choices = list("One Sample Variance Test" = "var_one", "Two Sample F-test" = "f_test", "Levene's Test" = "levene", "Bartlett's Test" = "bartlett")),
                        selectInput("var_test_var", "Variabel:", choices = NULL),
                        conditionalPanel("input.var_test_type == 'var_one'", numericInput("sigma_squared", "Nilai σ² Hipotesis:", value = 1, min = 0.01)),
                        conditionalPanel("input.var_test_type != 'var_one'", selectInput("var_group_var", "Variabel Grouping:", choices = NULL)),
                        sliderInput("var_alpha", "Tingkat Signifikansi:", min = 0.01, max = 0.1, value = 0.05, step = 0.01),
                        actionButton("run_var_test", "📈 Jalankan Uji", class = "btn-primary"),
                        br(), br(),
                        h4("💾 Download Hasil:"),
                        downloadButton("download_var_results", "📊 Hasil (Word)", class="btn-success"),
                        downloadButton("download_var_interpretation", "📝 Interpretasi (Word)", class="btn-info")
                      )
          ),
          argonColumn(width = 8,
                      argonCard(
                        width = 12,
                        status = "info",
                        title = "📈 Hasil Uji Varians",
                        argonTabSet(
                          id = "var_tabs",
                          width = 12,
                          argonTab(tabName = "Hasil Uji", withSpinner(verbatimTextOutput("var_test_results"))),
                          argonTab(
                            tabName = "Visualisasi", 
                            withSpinner(plotOutput("var_test_plot", height = "400px")),
                            br(),
                            downloadButton("download_var_plot_jpg", "🖼️ Unduh Plot (JPG)", class="btn-default")
                          ),
                          argonTab(tabName = "Interpretasi", withSpinner(verbatimTextOutput("var_test_interpretation")))
                        )
                      )
          )
        )
      ),
      # Halaman ANOVA
      argonTabItem(
        tabName = "anova",
        argonRow(
          argonColumn(width = 4,
                      argonCard(
                        width = 12,
                        status = "primary",
                        title = "📊 Pengaturan ANOVA",
                        radioButtons("anova_type", "Jenis ANOVA:", choices = list("One-Way ANOVA" = "one_way", "Two-Way ANOVA" = "two_way", "Repeated Measures ANOVA" = "repeated")),
                        selectInput("anova_response", "Variabel Response:", choices = NULL),
                        selectInput("anova_factor1", "Faktor 1:", choices = NULL),
                        conditionalPanel("input.anova_type == 'two_way'", selectInput("anova_factor2", "Faktor 2:", choices = NULL), checkboxInput("anova_interaction", "Sertakan Interaksi", TRUE)),
                        conditionalPanel("input.anova_type == 'repeated'", selectInput("anova_subject", "Variabel Subjek:", choices = NULL)),
                        checkboxInput("anova_posthoc", "Uji Post-hoc", TRUE),
                        conditionalPanel("input.anova_posthoc", selectInput("posthoc_method", "Metode Post-hoc:", choices = list("Tukey HSD" = "tukey", "Bonferroni" = "bonferroni", "Scheffe" = "scheffe", "LSD" = "lsd"))),
                        checkboxInput("anova_assumptions", "Uji Asumsi", TRUE),
                        sliderInput("anova_alpha", "Tingkat Signifikansi:", min = 0.01, max = 0.1, value = 0.05, step = 0.01),
                        actionButton("run_anova", "📊 Jalankan ANOVA", class = "btn-primary"),
                        br(), br(),
                        h4("💾 Download Hasil:"),
                        downloadButton("download_anova_results", "📊 Hasil (Word)", class="btn-success"),
                        downloadButton("download_anova_plots", "🖼️ Plot (JPG)", class="btn-warning"),
                        downloadButton("download_anova_interpretation", "📝 Interpretasi (Word)", class="btn-info")
                      )
          ),
          argonColumn(width = 8,
                      argonCard(
                        width = 12,
                        status = "info",
                        title = "📊 Hasil ANOVA",
                        argonTabSet(
                          id = "anova_tabs",
                          width = 12,
                          argonTab(tabName = "Tabel ANOVA", withSpinner(verbatimTextOutput("anova_results"))),
                          argonTab(tabName = "Visualisasi", withSpinner(plotOutput("anova_plot", height = "500px"))),
                          argonTab(tabName = "Post-hoc Tests", withSpinner(verbatimTextOutput("posthoc_results")), br(), withSpinner(plotOutput("posthoc_plot"))),
                          argonTab(tabName = "Uji Asumsi", withSpinner(verbatimTextOutput("anova_assumptions_results")), br(), withSpinner(plotOutput("anova_assumptions_plots"))),
                          argonTab(tabName = "Interpretasi", withSpinner(verbatimTextOutput("anova_interpretation")))
                        )
                      )
          )
        )
      ),
      # Halaman Regresi
      argonTabItem(
        tabName = "regression",
        argonRow(
          argonColumn(width = 4,
                      argonCard(
                        width = 12,
                        status = "primary",
                        title = "📈 Pengaturan Regresi Linear Berganda",
                        selectInput("reg_response", "Variabel Response:", choices = NULL),
                        selectInput("reg_predictors", "Variabel Prediktor:", choices = NULL, multiple = TRUE),
                        checkboxInput("reg_intercept", "Sertakan Intercept", TRUE),
                        checkboxInput("reg_interaction", "Sertakan Interaksi", FALSE),
                        checkboxInput("reg_polynomial", "Sertakan Polynomial", FALSE),
                        conditionalPanel("input.reg_polynomial", numericInput("poly_degree", "Derajat Polynomial:", value = 2, min = 2, max = 4)),
                        selectInput("reg_method", "Metode Seleksi:", choices = list("Enter (Forced)" = "enter", "Forward Selection" = "forward", "Backward Elimination" = "backward", "Stepwise" = "stepwise")),
                        checkboxInput("reg_standardize", "Standardisasi Variabel", FALSE),
                        checkboxInput("reg_diagnostics", "Uji Asumsi Lengkap", TRUE),
                        actionButton("run_regression", "📈 Jalankan Regresi", class = "btn-primary"),
                        br(), br(),
                        h4("💾 Download Hasil:"),
                        downloadButton("download_reg_results", "📊 Hasil (Word)", class="btn-success"),
                        downloadButton("download_reg_diagnostics", "🖼️ Diagnostik (JPG)", class="btn-warning"),
                        downloadButton("download_reg_interpretation", "📝 Interpretasi (Word)", class="btn-info")
                      )
          ),
          argonColumn(width = 8,
                      argonCard(
                        width = 12,
                        status = "info",
                        title = "📈 Hasil Regresi Linear Berganda",
                        argonTabSet(
                          id = "reg_tabs",
                          width = 12,
                          argonTab(tabName = "Summary Model", withSpinner(verbatimTextOutput("regression_summary"))),
                          argonTab(tabName = "Plot Diagnostik", withSpinner(plotOutput("regression_diagnostics", height = "600px"))),
                          argonTab(tabName = "Uji Asumsi", withSpinner(verbatimTextOutput("regression_assumptions")), br(), withSpinner(plotOutput("regression_assumptions_plots"))),
                          argonTab(tabName = "Perbandingan Model", withSpinner(DT::dataTableOutput("model_comparison")), br(), withSpinner(plotOutput("model_comparison_plot"))),
                          argonTab(tabName = "Prediksi", 
                                   argonRow(
                                     argonColumn(width=6, h4("Input Nilai untuk Prediksi:"), uiOutput("prediction_inputs")),
                                     argonColumn(width=6, h4("Hasil Prediksi:"), withSpinner(verbatimTextOutput("prediction_results")))
                                   ),
                                   br(),
                                   withSpinner(plotOutput("prediction_plot"))
                          ),
                          argonTab(tabName = "Interpretasi", withSpinner(verbatimTextOutput("regression_interpretation")))
                        )
                      )
          )
        )
      )
    )
  ),
  footer = argonDashFooter(
    copyrights = "@ STIS Komputasi Statistik, 2025",
    src = "https://www.stis.ac.id/",
    argonFooterMenu(
      argonFooterItem("Github", src = "https://github.com/bmlmcmc/naspaclust"),
      argonFooterItem("RInterface", src = "https://rinterface.github.io/argonDash/")
    )
  )
)


# Server Logic
server <- function(input, output, session) {
  
  # Reactive values untuk menyimpan semua hasil analisis
  values <- reactiveValues(
    categorized_data = NULL,
    current_plot = NULL,
    current_map = NULL,
    descriptive_results = NULL,
    correlation_results = NULL,
    distance_analysis_results = NULL,
    normality_results = NULL,
    homogeneity_results = NULL,
    mean_test_results = NULL,
    prop_test_results = NULL,
    var_test_results = NULL,
    anova_results = NULL,
    regression_model = NULL
  )
  
  # Inisialisasi semua pilihan input
  observe({
    req(sovi_data)
    
    # Identifikasi jenis variabel
    numeric_vars <- names(sovi_data)[sapply(sovi_data, is.numeric)]
    all_vars <- names(sovi_data)
    categorical_vars <- names(sovi_data)[sapply(sovi_data, function(x) {
      is.factor(x) || is.character(x) || 
        (is.numeric(x) && length(unique(x[!is.na(x)])) <= 10)
    })]
    
    # Jika tidak ada variabel kategorikal, buat dari numerik
    if(length(categorical_vars) == 0 && length(numeric_vars) >= 2) {
      var1 <- numeric_vars[1]
      var2 <- numeric_vars[2]
      
      sovi_data[[paste0(var1, "_Group")]] <<- cut(sovi_data[[var1]], 
                                                  breaks = 3, 
                                                  labels = c("Low", "Medium", "High"))
      sovi_data[[paste0(var2, "_Type")]] <<- cut(sovi_data[[var2]], 
                                                 breaks = 2, 
                                                 labels = c("Type_A", "Type_B"))
      
      categorical_vars <- c(categorical_vars, paste0(var1, "_Group"), paste0(var2, "_Type"))
    }
    
    # Update semua select inputs
    updateSelectInput(session, "cat_variable", choices = numeric_vars)
    updateCheckboxGroupInput(session, "desc_vars", choices = numeric_vars, selected = numeric_vars[1:min(5, length(numeric_vars))])
    updateSelectInput(session, "desc_group_var", choices = categorical_vars)
    updateSelectInput(session, "plot_x", choices = numeric_vars)
    updateSelectInput(session, "plot_y", choices = numeric_vars)
    updateSelectInput(session, "plot_group", choices = categorical_vars)
    updateSelectInput(session, "map_variable", choices = numeric_vars)
    updateSelectInput(session, "norm_var", choices = numeric_vars)
    updateSelectInput(session, "homo_var", choices = numeric_vars)
    updateSelectInput(session, "homo_group", choices = categorical_vars)
    updateSelectInput(session, "mean_test_var", choices = numeric_vars)
    updateSelectInput(session, "mean_group_var", choices = categorical_vars)
    updateSelectInput(session, "paired_var", choices = numeric_vars)
    updateSelectInput(session, "prop_var", choices = all_vars)
    updateSelectInput(session, "prop_group_var", choices = categorical_vars)
    updateSelectInput(session, "var_test_var", choices = numeric_vars)
    updateSelectInput(session, "var_group_var", choices = categorical_vars)
    updateSelectInput(session, "anova_response", choices = numeric_vars)
    updateSelectInput(session, "anova_factor1", choices = categorical_vars)
    updateSelectInput(session, "anova_factor2", choices = categorical_vars)
    updateSelectInput(session, "anova_subject", choices = all_vars)
    updateSelectInput(session, "reg_response", choices = numeric_vars)
    updateSelectInput(session, "reg_predictors", choices = numeric_vars)
  })
  
  # ===== OUTPUT BERANDA =====
  
  output$home_sovi_summary <- DT::renderDataTable({
    if(exists("sovi_data")) {
      summary_data <- data.frame(
        Variabel = names(sovi_data),
        Tipe = sapply(sovi_data, class),
        Missing = sapply(sovi_data, function(x) sum(is.na(x))),
        Min = sapply(sovi_data, function(x) if(is.numeric(x)) round(min(x, na.rm = TRUE), 3) else "-"),
        Max = sapply(sovi_data, function(x) if(is.numeric(x)) round(max(x, na.rm = TRUE), 3) else "-"),
        Mean = sapply(sovi_data, function(x) if(is.numeric(x)) round(mean(x, na.rm = TRUE), 3) else "-"),
        stringsAsFactors = FALSE
      )
      DT::datatable(summary_data, options = list(pageLength = 5, scrollX = TRUE), caption = "Ringkasan Statistik Dataset SoVI")
    }
  }, server = FALSE)
  
  output$home_quick_viz <- renderPlot({
    if(exists("sovi_data")) {
      numeric_vars <- names(sovi_data)[sapply(sovi_data, is.numeric)]
      if(length(numeric_vars) > 0) {
        var_name <- numeric_vars[1]
        ggplot(sovi_data, aes_string(x = var_name)) +
          geom_histogram(fill = "#5e72e4", alpha = 0.7, bins = 30) +
          theme_minimal() +
          labs(title = paste("Distribusi", var_name), x = var_name, y = "Frekuensi") +
          theme(plot.title = element_text(hjust = 0.5))
      }
    }
  })
  
  output$distance_info <- renderText({
    if(exists("distance_matrix")) {
      paste(
        "INFORMASI MATRIKS JARAK",
        "========================",
        "",
        paste("Dimensi Matriks:", nrow(distance_matrix), "x", ncol(distance_matrix)),
        paste("Tipe Data: Matriks Penimbang Jarak"),
        "",
        "Statistik Deskriptif Jarak:",
        paste("• Jarak Minimum:", round(min(distance_matrix, na.rm = TRUE), 3)),
        paste("• Jarak Maksimum:", round(max(distance_matrix, na.rm = TRUE), 3)),
        paste("• Jarak Rata-rata:", round(mean(as.matrix(distance_matrix), na.rm = TRUE), 3)),
        sep = "\n"
      )
    } else {
      "Data matriks jarak tidak tersedia."
    }
  })
  
  output$distance_stats_table <- DT::renderDataTable({
    if(exists("distance_matrix")) {
      stats_data <- data.frame(
        ID = 1:nrow(distance_matrix),
        Min_Distance = apply(distance_matrix, 1, function(x) round(min(x[x > 0], na.rm = TRUE), 3)),
        Max_Distance = apply(distance_matrix, 1, function(x) round(max(x, na.rm = TRUE), 3)),
        Mean_Distance = apply(distance_matrix, 1, function(x) round(mean(x, na.rm = TRUE), 3)),
        SD_Distance = apply(distance_matrix, 1, function(x) round(sd(x, na.rm = TRUE), 3))
      )
      DT::datatable(stats_data, options = list(pageLength = 5, scrollX = TRUE), caption = "Statistik Jarak per Observasi")
    }
  }, server = FALSE)
  
  # ===== MANAJEMEN DATA =====
  
  observeEvent(input$apply_categorization, {
    req(input$cat_variable, input$cat_groups)
    
    var_data <- sovi_data[[input$cat_variable]]
    
    if(input$cat_method == "quantile") {
      breaks <- quantile(var_data, probs = seq(0, 1, length.out = input$cat_groups + 1), na.rm = TRUE)
      categorized <- cut(var_data, breaks = breaks, include.lowest = TRUE,
                         labels = if(input$cat_labels != "") {
                           trimws(strsplit(input$cat_labels, ",")[[1]])
                         } else {
                           paste0("Q", 1:input$cat_groups)
                         })
    } else if(input$cat_method == "equal") {
      breaks <- seq(min(var_data, na.rm = TRUE), max(var_data, na.rm = TRUE), 
                    length.out = input$cat_groups + 1)
      categorized <- cut(var_data, breaks = breaks, include.lowest = TRUE,
                         labels = if(input$cat_labels != "") {
                           trimws(strsplit(input$cat_labels, ",")[[1]])
                         } else {
                           paste0("C", 1:input$cat_groups)
                         })
    } else if(input$cat_method == "kmeans") {
      km_result <- kmeans(var_data[!is.na(var_data)], centers = input$cat_groups)
      categorized <- rep(NA, length(var_data))
      categorized[!is.na(var_data)] <- paste0("K", km_result$cluster)
      categorized <- as.factor(categorized)
    }
    
    values$categorized_data <- sovi_data
    values$categorized_data[[paste0(input$cat_variable, "_categorized")]] <- categorized
    
    showNotification("✅ Kategorisasi berhasil diterapkan!", type = "message")
  })
  
  output$categorized_data <- DT::renderDataTable({
    if(is.null(values$categorized_data)) {
      return(DT::datatable(data.frame(Message = "Belum ada data terkategorisasi. Silakan terapkan kategorisasi terlebih dahulu.")))
    }
    DT::datatable(values$categorized_data, 
                  options = list(scrollX = TRUE, pageLength = 10),
                  class = 'cell-border stripe hover')
  })
  
  # Reactive expression for categorization plot
  categorization_plot_object <- reactive({
    if(is.null(values$categorized_data) || is.null(input$cat_variable)) {
      return(NULL)
    }
    
    cat_var <- paste0(input$cat_variable, "_categorized")
    
    p1 <- ggplot(values$categorized_data, aes_string(x = input$cat_variable)) +
      geom_histogram(bins = 30, fill = "#5e72e4", alpha = 0.7, color = "white") +
      theme_minimal() +
      labs(title = "Distribusi Asli", x = input$cat_variable, y = "Frekuensi")
    
    p2 <- ggplot(values$categorized_data, aes_string(x = cat_var)) +
      geom_bar(fill = "#11cdef", alpha = 0.7, color = "white") +
      theme_minimal() +
      labs(title = "Distribusi Terkategorisasi", x = "Kategori", y = "Frekuensi") +
      theme(axis.text.x = element_text(angle = 45, hjust = 1))
    
    grid.arrange(p1, p2, ncol = 2)
  })
  
  output$categorization_plot <- renderPlot({
    p <- categorization_plot_object()
    if(is.null(p)){
      plot.new()
      text(0.5, 0.5, "Belum ada data untuk divisualisasikan.\nSilakan terapkan kategorisasi terlebih dahulu.", 
           cex = 1.2, col = "gray")
    } else {
      p
    }
  })
  
  output$category_stats <- DT::renderDataTable({
    if(is.null(values$categorized_data) || is.null(input$cat_variable)) {
      return(DT::datatable(data.frame(Message = "Belum ada data terkategorisasi.")))
    }
    
    cat_var <- paste0(input$cat_variable, "_categorized")
    
    # Statistik per kategori
    cat_stats <- values$categorized_data %>%
      group_by_at(cat_var) %>%
      summarise(
        N = n(),
        Mean_Original = round(mean(get(input$cat_variable), na.rm = TRUE), 4),
        Median_Original = round(median(get(input$cat_variable), na.rm = TRUE), 4),
        SD_Original = round(sd(get(input$cat_variable), na.rm = TRUE), 4),
        Min_Original = round(min(get(input$cat_variable), na.rm = TRUE), 4),
        Max_Original = round(max(get(input$cat_variable), na.rm = TRUE), 4),
        .groups = 'drop'
      )
    
    DT::datatable(cat_stats, 
                  options = list(scrollX = TRUE, pageLength = 10),
                  class = 'cell-border stripe hover')
  })
  
  # Reactive expression for categorization interpretation text
  cat_interpretation_reactive <- reactive({
    if(is.null(values$categorized_data) || is.null(input$cat_variable)) {
      return("Belum ada data terkategorisasi. Silakan terapkan kategorisasi terlebih dahulu.")
    }
    
    cat_var <- paste0(input$cat_variable, "_categorized")
    freq_table <- table(values$categorized_data[[cat_var]])
    
    method_desc <- switch(input$cat_method,
                          "quantile" = "berdasarkan kuantil (pembagian sama banyak)",
                          "equal" = "berdasarkan range yang sama",
                          "kmeans" = "berdasarkan clustering K-means"
    )
    
    paste0(
      "📊 INTERPRETASI KATEGORISASI DATA\n",
      "=================================\n\n",
      "🎯 Informasi Kategorisasi:\n",
      "• Variabel: ", input$cat_variable, "\n",
      "• Metode: ", str_to_title(method_desc), "\n",
      "• Jumlah Kategori: ", input$cat_groups, "\n\n",
      "📈 Distribusi Kategori:\n",
      paste(names(freq_table), ":", freq_table, "observasi (", 
            round(freq_table/sum(freq_table)*100, 1), "%)", collapse = "\n"), "\n\n",
      "🔍 Interpretasi Statistik:\n",
      "Kategorisasi ini mengubah variabel kontinyu menjadi kategorikal untuk:\n",
      "• Memudahkan analisis grup dan perbandingan\n",
      "• Mengurangi pengaruh outlier dalam data\n",
      "• Memungkinkan penggunaan uji non-parametrik\n",
      "• Interpretasi yang lebih mudah dipahami\n\n",
      "⚠️ Catatan Penting:\n",
      "• Kategorisasi mengurangi informasi detail dari data asli\n",
      "• Pilihan jumlah kategori mempengaruhi hasil analisis\n"
    )
  })
  
  output$categorization_interpretation <- renderText({
    cat_interpretation_reactive()
  })
  
  # ===== STATISTIK DESKRIPTIF =====
  
  observeEvent(input$calc_descriptive, {
    req(input$desc_vars)
    
    # Define the list of summary functions
    summary_funs <- list(
      N = ~sum(!is.na(.)),
      Mean = ~round(mean(., na.rm = TRUE), 4),
      Median = ~round(median(., na.rm = TRUE), 4),
      SD = ~round(sd(., na.rm = TRUE), 4),
      Variance = ~round(var(., na.rm = TRUE), 4),
      Min = ~round(min(., na.rm = TRUE), 4),
      Max = ~round(max(., na.rm = TRUE), 4),
      Q1 = ~round(quantile(., 0.25, na.rm = TRUE), 4),
      Q3 = ~round(quantile(., 0.75, na.rm = TRUE), 4),
      IQR = ~round(IQR(., na.rm = TRUE), 4)
    )
    
    if(input$desc_by_group && !is.null(input$desc_group_var) && input$desc_group_var != "") {
      # Logic for grouped analysis using modern dplyr
      desc_results <- sovi_data %>%
        group_by(!!sym(input$desc_group_var)) %>%
        summarise(across(all_of(input$desc_vars), summary_funs), .groups = 'drop')
    } else {
      # Logic for ungrouped analysis using modern dplyr
      desc_results <- sovi_data %>%
        summarise(across(all_of(input$desc_vars), summary_funs)) %>%
        pivot_longer(
          cols = everything(),
          names_to = c("Variable", ".value"),
          names_sep = "_"
        )
    }
    
    values$descriptive_results <- desc_results
    showNotification("✅ Statistik deskriptif berhasil dihitung!", type = "message")
  })
  
  output$descriptive_stats <- DT::renderDataTable({
    if(is.null(values$descriptive_results)) {
      return(DT::datatable(data.frame(Message = "Belum ada hasil statistik deskriptif. Silakan hitung terlebih dahulu.")))
    }
    DT::datatable(values$descriptive_results, 
                  options = list(scrollX = TRUE, pageLength = 15),
                  class = 'cell-border stripe hover') %>%
      DT::formatRound(columns = sapply(values$descriptive_results, is.numeric), digits = 4)
  })
  
  # Reactive expression for descriptive plots
  descriptive_plots_object <- reactive({
    req(input$desc_vars)
    
    plot_list <- lapply(input$desc_vars, function(var_name) {
      p1 <- ggplot(sovi_data, aes_string(x = var_name)) +
        geom_histogram(bins = 30, fill = "#5e72e4", alpha = 0.7, color = "white") +
        theme_minimal() +
        labs(title = paste("Histogram:", var_name), x = var_name, y = "Frekuensi")
      
      p2 <- ggplot(sovi_data, aes_string(y = var_name)) +
        geom_boxplot(fill = "#11cdef", alpha = 0.7) +
        theme_minimal() +
        labs(title = paste("Boxplot:", var_name), y = var_name) +
        theme(axis.text.x = element_blank(), axis.ticks.x = element_blank())
      
      list(p1, p2)
    })
    
    # Flatten the list of plots
    plot_list <- unlist(plot_list, recursive = FALSE)
    
    # Arrange plots in a grid
    marrangeGrob(plot_list, ncol = 2, nrow = length(input$desc_vars))
  })
  
  output$descriptive_plots <- renderPlot({
    descriptive_plots_object()
  })
  
  # Reactive expression for group comparison plot
  group_comparison_plot_object <- reactive({
    if(!input$desc_by_group || is.null(input$desc_group_var) || is.null(input$desc_vars)) {
      return(NULL)
    }
    
    var_name <- input$desc_vars[1]
    group_var <- input$desc_group_var
    
    ggplot(sovi_data, aes_string(x = group_var, y = var_name, fill = group_var)) +
      geom_boxplot(alpha = 0.7) +
      geom_jitter(width = 0.2, alpha = 0.5) +
      scale_fill_viridis_d() +
      theme_minimal() +
      labs(title = paste("Perbandingan", var_name, "berdasarkan", group_var),
           x = group_var, y = var_name) +
      theme(legend.position = "none")
  })
  
  output$group_comparison_plot <- renderPlot({
    p <- group_comparison_plot_object()
    if(is.null(p)){
      plot.new()
      text(0.5, 0.5, "Aktifkan 'Analisis per Grup' untuk melihat perbandingan.", cex = 1.2, col = "gray")
    } else {
      p
    }
  })
  
  output$group_stats_table <- DT::renderDataTable({
    if(!input$desc_by_group || is.null(input$desc_group_var) || is.null(input$desc_vars)) {
      return(DT::datatable(data.frame(Message = "Aktifkan 'Analisis per Grup' untuk melihat statistik per grup.")))
    }
    
    group_stats <- sovi_data %>%
      select(all_of(c(input$desc_vars, input$desc_group_var))) %>%
      group_by_at(input$desc_group_var) %>%
      summarise_if(is.numeric, list(
        N = ~sum(!is.na(.)),
        Mean = ~round(mean(., na.rm = TRUE), 3),
        SD = ~round(sd(., na.rm = TRUE), 3)
      ), .groups = 'drop')
    
    DT::datatable(group_stats, 
                  options = list(scrollX = TRUE, pageLength = 10),
                  class = 'cell-border stripe hover')
  })
  
  # Reactive expression for descriptive interpretation text
  desc_interpretation_reactive <- reactive({
    if(is.null(values$descriptive_results)) {
      return("Belum ada hasil statistik deskriptif. Silakan hitung terlebih dahulu.")
    }
    
    paste0(
      "📊 INTERPRETASI STATISTIK DESKRIPTIF\n",
      "====================================\n\n",
      "🎯 Variabel yang Dianalisis:\n",
      paste("•", input$desc_vars, collapse = "\n"), "\n\n",
      "📈 Interpretasi Ukuran Pemusatan:\n",
      "• Mean (Rata-rata): Nilai tengah aritmatika dari data\n",
      "• Median: Nilai tengah yang membagi data menjadi dua bagian sama\n\n",
      "📊 Interpretasi Ukuran Penyebaran:\n",
      "• SD (Standar Deviasi): Ukuran penyebaran data dari rata-rata\n",
      "• IQR (Interquartile Range): Rentang 50% data tengah\n\n",
      "💡 Rekomendasi Analisis Lanjutan:\n",
      "• Jika distribusi normal: Gunakan uji parametrik\n",
      "• Jika distribusi tidak normal: Pertimbangkan transformasi atau uji non-parametrik\n"
    )
  })
  
  output$descriptive_interpretation <- renderText({
    desc_interpretation_reactive()
  })
  
  # ===== VISUALISASI & PETA =====
  
  observeEvent(input$generate_plot, {
    req(input$plot_x, input$plot_type)
    
    plot_title <- if(input$plot_title != "") input$plot_title else paste(str_to_title(input$plot_type), "of", input$plot_x)
    
    p <- ggplot(sovi_data)
    
    if(input$plot_type == "scatter" && !is.null(input$plot_y)) {
      p <- p + aes_string(x = input$plot_x, y = input$plot_y) +
        geom_point(alpha = 0.6, color = input$plot_color)
      if(input$add_trend) {
        p <- p + geom_smooth(method = "lm", se = TRUE, color = "#f5365c")
      }
    } else if(input$plot_type == "line" && !is.null(input$plot_y)) {
      p <- p + aes_string(x = input$plot_x, y = input$plot_y) +
        geom_line(color = input$plot_color, size = 1) +
        geom_point(color = input$plot_color, alpha = 0.7)
    } else if(input$plot_type == "hist") {
      p <- p + aes_string(x = input$plot_x) +
        geom_histogram(bins = 30, fill = input$plot_color, alpha = 0.7, color = "white")
    } else if(input$plot_type == "density") {
      p <- p + aes_string(x = input$plot_x) +
        geom_density(fill = input$plot_color, alpha = 0.7)
    } else if(input$plot_type == "box" && !is.null(input$plot_group)) {
      p <- p + aes_string(x = input$plot_group, y = input$plot_x) +
        geom_boxplot(fill = input$plot_color, alpha = 0.7)
    } else if(input$plot_type == "violin" && !is.null(input$plot_group)) {
      p <- p + aes_string(x = input$plot_group, y = input$plot_x) +
        geom_violin(fill = input$plot_color, alpha = 0.7) +
        geom_boxplot(width = 0.1, fill = "white", alpha = 0.8)
    } else if(input$plot_type == "heatmap") {
      numeric_data <- sovi_data %>% select_if(is.numeric)
      cor_matrix <- cor(numeric_data, use = "complete.obs")
      cor_long <- reshape2::melt(cor_matrix)
      p <- ggplot(cor_long, aes(Var1, Var2, fill = value)) +
        geom_tile() +
        scale_fill_gradient2(low = "#f5365c", high = "#5e72e4", mid = "white", midpoint = 0, limit = c(-1,1), name="Korelasi")
    }
    
    p <- p + theme_minimal() + labs(title = plot_title)
    
    values$current_plot <- p
    showNotification("✅ Plot berhasil dibuat!", type = "message")
  })
  
  output$interactive_plot <- renderPlotly({
    if(is.null(values$current_plot)) {
      plotly_empty() %>% layout(title = "Belum ada plot yang dibuat.")
    } else {
      ggplotly(values$current_plot)
    }
  })
  
  output$static_plot <- renderPlot({
    if(is.null(values$current_plot)) {
      plot.new()
      text(0.5, 0.5, "Belum ada plot yang dibuat.", cex = 1.2, col = "gray")
    } else {
      values$current_plot
    }
  })
  
  # ===== ANALISIS MATRIKS JARAK =====
  
  observeEvent(input$run_distance_analysis, {
    req(input$distance_analysis_type)
    
    if(input$distance_analysis_type == "descriptive") {
      distance_values <- as.matrix(distance_matrix)
      distance_values <- distance_values[upper.tri(distance_values)]
      desc_stats <- data.frame(
        Statistik = c("N", "Mean", "Median", "SD", "Min", "Max"),
        Nilai = c(length(distance_values), round(mean(distance_values, na.rm = TRUE), 4), round(median(distance_values, na.rm = TRUE), 4), round(sd(distance_values, na.rm = TRUE), 4), round(min(distance_values, na.rm = TRUE), 4), round(max(distance_values, na.rm = TRUE), 4))
      )
      values$distance_analysis_results <- list(type = "descriptive", stats = desc_stats, values = distance_values)
      
    } else if(input$distance_analysis_type == "clustering") {
      distance_values <- as.matrix(distance_matrix)
      if(input$cluster_method == "kmeans") {
        mds_result <- cmdscale(distance_values, k = 2)
        cluster_result <- kmeans(mds_result, centers = input$n_clusters)
        clusters <- cluster_result$cluster
      } else if(input$cluster_method == "hierarchical") {
        hc_result <- hclust(as.dist(distance_values))
        clusters <- cutree(hc_result, k = input$n_clusters)
      } else if(input$cluster_method == "pam") {
        pam_result <- cluster::pam(as.dist(distance_values), k = input$n_clusters)
        clusters <- pam_result$clustering
      }
      cluster_summary <- data.frame(Cluster = 1:input$n_clusters, N = as.vector(table(clusters)))
      values$distance_analysis_results <- list(type = "clustering", method = input$cluster_method, clusters = clusters, summary = cluster_summary, n_clusters = input$n_clusters)
      
    } else if(input$distance_analysis_type == "heatmap") {
      values$distance_analysis_results <- list(type = "heatmap", matrix = as.matrix(distance_matrix))
      
    } else if(input$distance_analysis_type == "pca") {
      distance_values <- as.matrix(distance_matrix)
      mds_result <- cmdscale(distance_values, k = min(5, nrow(distance_values)-1), eig = TRUE)
      eigenvalues <- mds_result$eig[mds_result$eig > 0]
      prop_var <- eigenvalues / sum(eigenvalues)
      pca_summary <- data.frame(Component = 1:length(prop_var), Eigenvalue = round(eigenvalues, 4), Proportion = round(prop_var, 4), Cumulative = round(cumsum(prop_var), 4))
      values$distance_analysis_results <- list(type = "pca", coordinates = mds_result$points, summary = pca_summary, eigenvalues = eigenvalues)
    }
    
    showNotification("✅ Analisis matriks jarak selesai!", type = "message")
  })
  
  output$distance_analysis_results <- renderPrint({
    if(is.null(values$distance_analysis_results)) return("Belum ada hasil analisis.")
    result <- values$distance_analysis_results
    cat("📏 HASIL ANALISIS MATRIKS JARAK\n")
    if(result$type == "descriptive") {
      print(result$stats)
    } else if(result$type == "clustering") {
      cat("Metode:", toupper(result$method), "\nJumlah Cluster:", result$n_clusters, "\n\nRingkasan Cluster:\n")
      print(result$summary)
    } else if(result$type == "pca") {
      cat("Ringkasan Komponen:\n")
      print(result$summary)
    }
  })
  
  output$distance_results_table <- DT::renderDataTable({
    # Pastikan hasil analisis sudah ada
    if(is.null(values$distance_analysis_results)) {
      return(DT::datatable(data.frame(Message = "Belum ada hasil analisis untuk ditampilkan.")))
    }
    
    result <- values$distance_analysis_results
    
    # Siapkan data frame yang akan ditampilkan berdasarkan jenis analisis
    table_data <- NULL
    if (result$type == "descriptive") {
      table_data <- result$stats
    } else if (result$type == "clustering") {
      table_data <- result$summary # Ini adalah data frame ringkasan cluster
    } else if (result$type == "pca") {
      table_data <- result$summary # Ini adalah data frame ringkasan PCA/MDS
    }
    
    # Tampilkan tabel hanya jika table_data adalah data frame atau matriks
    if (!is.null(table_data) && (is.data.frame(table_data) || is.matrix(table_data))) {
      DT::datatable(table_data, 
                    options = list(pageLength = 10, scrollX = TRUE), 
                    caption = "Tabel Hasil Analisis Matriks Jarak")
    } else {
      DT::datatable(data.frame(Message = "Tidak ada data dalam bentuk tabel untuk jenis analisis ini."))
    }
  })
  
  output$distance_plot <- renderPlot({
    if(is.null(values$distance_analysis_results)) {
      plot.new(); text(0.5, 0.5, "Belum ada hasil analisis.", cex = 1.2, col = "gray")
      return()
    }
    result <- values$distance_analysis_results
    if(result$type == "descriptive") {
      hist(result$values, breaks = 30, col = "#5e72e4", main = "Distribusi Jarak", xlab = "Jarak")
    } else if(result$type == "clustering") {
      mds_coords <- cmdscale(as.matrix(distance_matrix), k = 2)
      plot_data <- data.frame(X = mds_coords[,1], Y = mds_coords[,2], Cluster = as.factor(result$clusters))
      ggplot(plot_data, aes(x = X, y = Y, color = Cluster)) +
        geom_point(size = 3, alpha = 0.7) +
        scale_color_viridis_d() +
        theme_minimal() +
        labs(title = paste("Hasil Clustering -", str_to_title(result$method)))
    } else if(result$type == "pca") {
      plot_data <- data.frame(Component = 1:length(result$eigenvalues), Eigenvalue = result$eigenvalues)
      ggplot(plot_data, aes(x = Component, y = Eigenvalue)) +
        geom_line(color = "#5e72e4", size = 1) + geom_point(color = "#11cdef", size = 3) +
        theme_minimal() + labs(title = "Scree Plot - Analisis Komponen Utama")
    }
  })
  
  output$distance_heatmap <- renderPlot({
    if(is.null(values$distance_analysis_results)) {
      plot.new(); text(0.5, 0.5, "Belum ada hasil analisis.", cex = 1.2, col = "gray")
      return()
    }
    distance_matrix_num <- as.matrix(distance_matrix)
    if(nrow(distance_matrix_num) > 100) {
      indices <- sample(1:nrow(distance_matrix_num), 100)
      distance_matrix_num <- distance_matrix_num[indices, indices]
    }
    heatmap(distance_matrix_num, symm = TRUE, main = "Heatmap Matriks Jarak")
  })
  
  output$distance_interpretation <- renderText({
    if(is.null(values$distance_analysis_results)) return("Belum ada hasil analisis.")
    result <- values$distance_analysis_results
    if(result$type == "descriptive") {
      "INTERPRETASI: Matriks jarak menggambarkan kedekatan antar observasi. Jarak rata-rata menunjukkan tingkat heterogenitas data."
    } else if(result$type == "clustering") {
      "INTERPRETASI: Cluster mengelompokkan observasi berdasarkan kedekatan jarak. Ukuran cluster menunjukkan distribusi data."
    } else if(result$type == "pca") {
      "INTERPRETASI: MDS digunakan untuk visualisasi matriks jarak. Dua komponen pertama menjelaskan proporsi varians terbesar."
    } else {
      "Interpretasi akan muncul setelah analisis selesai."
    }
  })
  
  # ===== UJI ASUMSI =====
  
  observeEvent(input$run_normality, {
    req(input$norm_var, input$norm_tests)
    var_data <- sovi_data[[input$norm_var]]; var_data <- var_data[!is.na(var_data)]
    test_results <- list()
    if("shapiro" %in% input$norm_tests) test_results$shapiro <- shapiro.test(if(length(var_data)>5000) sample(var_data, 5000) else var_data)
    if("ks" %in% input$norm_tests) test_results$ks <- ks.test(var_data, "pnorm", mean(var_data), sd(var_data))
    if("lillie" %in% input$norm_tests) test_results$lillie <- lillie.test(var_data)
    if("ad" %in% input$norm_tests) test_results$ad <- ad.test(var_data)
    if("jb" %in% input$norm_tests && requireNamespace("tseries", quietly = TRUE)) test_results$jb <- tseries::jarque.bera.test(var_data)
    values$normality_results <- list(tests = test_results, variable = input$norm_var, data = var_data, alpha = input$norm_alpha)
    showNotification("✅ Uji normalitas selesai!", type = "message")
  })
  
  output$normality_results <- renderPrint({
    if(is.null(values$normality_results)) return("Belum ada hasil uji normalitas.")
    cat("🔍 HASIL UJI NORMALITAS\nVariabel:", values$normality_results$variable, "\nα =", values$normality_results$alpha, "\n\n")
    for(test_name in names(values$normality_results$tests)) {
      test_result <- values$normality_results$tests[[test_name]]
      cat(str_to_upper(test_name), "TEST:\n")
      cat("p-value =", format(test_result$p.value, scientific = TRUE), "\n")
      cat("Kesimpulan:", ifelse(test_result$p.value > values$normality_results$alpha, "✅ Normal", "❌ Tidak Normal"), "\n\n")
    }
  })
  
  output$normality_plots <- renderPlot({
    if(is.null(values$normality_results)) {
      plot.new(); text(0.5, 0.5, "Belum ada hasil uji normalitas.", cex = 1.2, col = "gray")
      return()
    }
    var_data <- values$normality_results$data
    par(mfrow = c(1, 2))
    hist(var_data, freq = FALSE, main = paste("Histogram:", values$normality_results$variable), xlab = values$normality_results$variable, col = "#5e72e4")
    curve(dnorm(x, mean = mean(var_data), sd = sd(var_data)), add = TRUE, col = "#f5365c", lwd = 2)
    qqnorm(var_data, main = "Q-Q Plot Normal", col = "#5e72e4"); qqline(var_data, col = "#f5365c", lwd = 2)
  })
  
  observeEvent(input$run_homogeneity, {
    req(input$homo_var, input$homo_group)
    group_var <- as.factor(sovi_data[[input$homo_group]])
    test_result <- switch(input$homo_test,
                          "levene" = car::leveneTest(sovi_data[[input$homo_var]], group_var),
                          "bartlett" = bartlett.test(sovi_data[[input$homo_var]], group_var),
                          "fligner" = fligner.test(sovi_data[[input$homo_var]], group_var))
    values$homogeneity_results <- list(test = test_result, variable = input$homo_var, group = input$homo_group, method = input$homo_test, alpha = input$homo_alpha, group_data = group_var)
    showNotification("✅ Uji homogenitas selesai!", type = "message")
  })
  
  output$homogeneity_results <- renderPrint({
    if(is.null(values$homogeneity_results)) return("Belum ada hasil uji homogenitas.")
    cat("⚖️ HASIL UJI HOMOGENITAS\nMetode:", str_to_upper(values$homogeneity_results$method), "\nα =", values$homogeneity_results$alpha, "\n\n")
    print(values$homogeneity_results$test)
    p_val <- if(values$homogeneity_results$method == "levene") values$homogeneity_results$test$`Pr(>F)`[1] else values$homogeneity_results$test$p.value
    cat("\nKesimpulan:", ifelse(p_val > values$homogeneity_results$alpha, "✅ Varians homogen", "❌ Varians tidak homogen"), "\n")
  })
  
  output$homogeneity_plots <- renderPlot({
    if(is.null(values$homogeneity_results)) {
      plot.new(); text(0.5, 0.5, "Belum ada hasil uji homogenitas.", cex = 1.2, col = "gray")
      return()
    }
    plot_data <- data.frame(value = sovi_data[[values$homogeneity_results$variable]], group = values$homogeneity_results$group_data)
    plot_data <- plot_data[complete.cases(plot_data), ]
    boxplot(value ~ group, data = plot_data, main = "Boxplot per Grup", xlab = values$homogeneity_results$group, ylab = values$homogeneity_results$variable, col = viridis(length(unique(plot_data$group))))
  })
  
  assumptions_interpretation_reactive <- reactive({
    "🎯 Pentingnya Uji Asumsi:\nUji asumsi diperlukan sebelum melakukan analisis parametrik untuk memastikan\nvaliditas dan reliabilitas hasil analisis.\n\n🔍 Uji Normalitas:\n• Menguji apakah data mengikuti distribusi normal\n• H0: Data berdistribusi normal\n• Jika p-value > α: Terima H0 (data normal)\n• Penting untuk: t-test, ANOVA, regresi linear\n\n⚖️ Uji Homogenitas:\n• Menguji apakah varians antar grup sama\n• H0: Varians antar grup sama (homogen)\n• Jika p-value > α: Terima H0 (varians homogen)\n• Penting untuk: ANOVA, t-test dua sampel\n\n📊 Kombinasi Hasil:\n• Normal + Homogen: Gunakan uji parametrik standar\n• Normal + Tidak Homogen: Gunakan Welch's test\n• Tidak Normal + Homogen: Pertimbangkan transformasi\n• Tidak Normal + Tidak Homogen: Gunakan uji non-parametrik\n\n💡 Rekomendasi Umum:\n• Selalu periksa asumsi sebelum analisis utama\n• Gunakan visualisasi untuk konfirmasi\n• Pertimbangkan ukuran sampel dalam interpretasi\n• Dokumentasikan hasil uji asumsi dalam laporan"
  })
  
  output$assumptions_interpretation <- renderText({
    assumptions_interpretation_reactive()
  })
  
  # ===== UJI RATA-RATA =====
  
  observeEvent(input$run_mean_test, {
    req(input$mean_test_type, input$mean_test_var)
    test_result <- NULL
    if(input$mean_test_type == "t_one") {
      test_result <- t.test(sovi_data[[input$mean_test_var]], mu = input$mu_value, alternative = input$mean_alternative, conf.level = 1 - input$mean_alpha)
    } else if(input$mean_test_type %in% c("t_two", "welch")) {
      req(input$mean_group_var)
      group_var <- as.factor(sovi_data[[input$mean_group_var]])
      if(length(levels(group_var)) > 2) group_var <- factor(group_var, levels = levels(group_var)[1:2])
      test_result <- t.test(sovi_data[[input$mean_test_var]] ~ group_var, alternative = input$mean_alternative, var.equal = (input$mean_test_type == "t_two"), conf.level = 1 - input$mean_alpha)
    } else if(input$mean_test_type == "t_paired") {
      req(input$paired_var)
      test_result <- t.test(sovi_data[[input$mean_test_var]], sovi_data[[input$paired_var]], paired = TRUE, alternative = input$mean_alternative, conf.level = 1 - input$mean_alpha)
    }
    if(!is.null(test_result)) {
      values$mean_test_results <- list(test = test_result, type = input$mean_test_type, variable = input$mean_test_var, group_var = input$mean_group_var, alpha = input$mean_alpha, alternative = input$mean_alternative)
      showNotification("✅ Uji rata-rata selesai!", type = "message")
    }
  })
  
  mean_results_text_reactive <- reactive({
    if(is.null(values$mean_test_results)) return("Belum ada hasil uji rata-rata.")
    
    hasil <- values$mean_test_results
    p_val <- hasil$test$p.value
    
    # Capture the original print output of the t.test object
    test_output <- capture.output(print(hasil$test))
    
    # Build interpretation text
    interp_text <- if(hasil$type == "t_one") {
      paste("H0: μ =", input$mu_value, "\nH1: μ", switch(hasil$alternative, "two.sided" = "≠", "greater" = ">", "less" = "<"), input$mu_value)
    } else {
      paste("H0: μ1 = μ2\nH1: μ1", switch(hasil$alternative, "two.sided" = "≠", "greater" = ">", "less" = "<"), "μ2")
    }
    
    keputusan <- ifelse(p_val > hasil$alpha, "Gagal menolak H0", "Tolak H0")
    kesimpulan <- ifelse(p_val > hasil$alpha, "Tidak ada bukti yang cukup untuk menyatakan perbedaan rata-rata", "Ada bukti yang cukup untuk menyatakan perbedaan rata-rata")
    ci <- hasil$test$conf.int
    ci_text <- paste0("Interval Kepercayaan ", (1-hasil$alpha)*100, "%: [", round(ci[1], 4), ", ", round(ci[2], 4), "]")
    
    # Combine all parts
    paste(
      "🧪 HASIL UJI RATA-RATA\n======================\n",
      paste("Jenis Uji:", str_to_upper(gsub("_", " ", hasil$type))),
      paste("Variabel:", hasil$variable),
      paste("Hipotesis Alternatif:", hasil$alternative),
      paste("Tingkat Signifikansi:", hasil$alpha),
      "\n",
      paste(test_output, collapse = "\n"),
      "\n\n📊 INTERPRETASI:",
      interp_text,
      paste("\nKeputusan:", keputusan),
      paste("📈 Kesimpulan:", kesimpulan),
      paste("\n", ci_text),
      sep = "\n"
    )
  })
  
  output$mean_test_results <- renderText({
    mean_results_text_reactive()
  })
  
  # Reactive untuk plot uji rata-rata
  mean_test_plot_object <- reactive({
    if(is.null(values$mean_test_results)) return(NULL)
    
    hasil <- values$mean_test_results
    
    if(hasil$type %in% c("t_one", "t_paired")) {
      var_data <- sovi_data[[hasil$variable]]
      p <- ggplot(data.frame(x=var_data), aes(x=x)) +
        geom_histogram(aes(y=..density..), bins=30, fill="#5e72e4", alpha=0.7) +
        geom_density(color="#f5365c", size=1) +
        geom_vline(xintercept = mean(var_data, na.rm=TRUE), color="blue", linetype="dashed", size=1) +
        labs(title=paste("Distribusi", hasil$variable), x=hasil$variable) +
        theme_minimal()
      if(hasil$type == "t_one") {
        p <- p + geom_vline(xintercept = input$mu_value, color="green", linetype="dotted", size=1)
      }
      return(p)
    } else {
      group_var <- as.factor(sovi_data[[hasil$group_var]])
      df_plot <- data.frame(value = sovi_data[[hasil$variable]], group = group_var)
      
      p <- ggplot(df_plot, aes(x=group, y=value, fill=group)) +
        geom_boxplot(alpha=0.7) +
        scale_fill_viridis_d() +
        labs(title=paste("Perbandingan", hasil$variable, "berdasarkan", hasil$group_var), x=hasil$group_var, y=hasil$variable) +
        theme_minimal() +
        theme(legend.position = "none")
      return(p)
    }
  })
  
  output$mean_test_plot <- renderPlot({
    p <- mean_test_plot_object()
    if(is.null(p)) {
      plot.new(); text(0.5, 0.5, "Belum ada hasil uji rata-rata.", cex = 1.2, col = "gray")
    } else {
      print(p)
    }
  })
  
  mean_interpretation_text_reactive <- reactive({
    if(is.null(values$mean_test_results)) return("Belum ada hasil uji rata-rata.")
    
    hasil <- values$mean_test_results
    test_desc <- switch(hasil$type,
                        "t_one" = "menguji apakah rata-rata populasi sama dengan nilai tertentu",
                        "t_two" = "menguji apakah rata-rata dua grup independen berbeda",
                        "t_paired" = "menguji apakah rata-rata dua pengukuran berpasangan berbeda",
                        "welch" = "menguji apakah rata-rata dua grup berbeda (varians tidak sama)"
    )
    
    effect_size_text <- ""
    if(hasil$type %in% c("t_two", "welch")) {
      group_var <- as.factor(sovi_data[[hasil$group_var]])
      levels_group <- levels(group_var)
      
      group1_data <- sovi_data[[hasil$variable]][group_var == levels_group[1]]
      group2_data <- sovi_data[[hasil$variable]][group_var == levels_group[2]]
      
      n1 <- length(na.omit(group1_data))
      n2 <- length(na.omit(group2_data))
      sd1 <- sd(group1_data, na.rm=TRUE)
      sd2 <- sd(group2_data, na.rm=TRUE)
      
      pooled_sd <- sqrt(((n1-1)*sd1^2 + (n2-1)*sd2^2) / (n1+n2-2))
      cohens_d <- abs(mean(group1_data, na.rm=TRUE) - mean(group2_data, na.rm=TRUE)) / pooled_sd
      
      effect_interpretation <- if(cohens_d < 0.2) "kecil" else if(cohens_d < 0.5) "sedang" else if(cohens_d < 0.8) "besar" else "sangat besar"
      
      effect_size_text <- paste0(
        "\n\n📏 UKURAN EFEK (COHEN'S D):\n",
        "Cohen's d = ", round(cohens_d, 3), "\n",
        "Interpretasi: Efek ", effect_interpretation, "\n",
        "• d < 0.2: Efek kecil\n",
        "• 0.2 ≤ d < 0.5: Efek sedang\n",
        "• 0.5 ≤ d < 0.8: Efek besar\n",
        "• d ≥ 0.8: Efek sangat besar"
      )
    }
    
    paste(
      "📝 INTERPRETASI UJI RATA-RATA\n",
      "=============================\n\n",
      "🎯 Tujuan Uji:\n",
      "Uji ini ", test_desc, ".\n\n",
      "📊 Hasil Statistik:\n",
      "• t-statistic = ", round(hasil$test$statistic, 4), "\n",
      "• df = ", hasil$test$parameter, "\n",
      "• p-value = ", hasil$test$p.value, "\n",
      "• α = ", hasil$alpha, "\n\n",
      "🔍 Interpretasi Praktis:\n",
      ifelse(hasil$test$p.value > hasil$alpha,
             "Tidak ada perbedaan yang signifikan secara statistik.",
             "Ada perbedaan yang signifikan secara statistik."), "\n\n",
      "📈 Interval Kepercayaan:\n",
      "Dengan tingkat kepercayaan ", (1-hasil$alpha)*100, "%, interval untuk perbedaan rata-rata adalah [",
      round(hasil$test$conf.int[1], 4), ", ",
      round(hasil$test$conf.int[2], 4), "].\n",
      "Jika interval mencakup 0, maka tidak ada perbedaan yang signifikan.",
      effect_size_text,
      "\n\n💡 Rekomendasi:\n",
      "• Periksa asumsi normalitas dan homogenitas sebelum interpretasi final\n",
      "• Pertimbangkan signifikansi praktis selain signifikansi statistik\n",
      "• Laporkan ukuran efek untuk memberikan konteks hasil",
      sep=""
    )
  })
  
  output$mean_test_interpretation <- renderText({
    mean_interpretation_text_reactive()
  })
  
  # ===== UJI PROPORSI =====
  
  observeEvent(input$run_prop_test, {
    req(input$prop_test_type, input$prop_var)
    test_result <- NULL
    if(input$prop_test_type == "prop_one") {
      var_data <- as.factor(sovi_data[[input$prop_var]])
      success_count <- sum(var_data == levels(var_data)[1], na.rm = TRUE)
      total_count <- sum(!is.na(var_data))
      test_result <- prop.test(success_count, total_count, p = input$prop_value)
    } else if(input$prop_test_type == "prop_two") {
      req(input$prop_group_var)
      cont_table <- table(sovi_data[[input$prop_var]], sovi_data[[input$prop_group_var]])
      test_result <- prop.test(cont_table)
    } else if(input$prop_test_type == "chisq_gof") {
      test_result <- chisq.test(table(sovi_data[[input$prop_var]]))
    } else if(input$prop_test_type == "chisq_indep") {
      req(input$prop_group_var)
      test_result <- chisq.test(table(sovi_data[[input$prop_var]], sovi_data[[input$prop_group_var]]))
    }
    if(!is.null(test_result)) {
      values$prop_test_results <- list(test = test_result, type = input$prop_test_type, variable = input$prop_var, group_var = input$prop_group_var, alpha = input$prop_alpha)
      showNotification("✅ Uji proporsi selesai!", type = "message")
    }
  })
  
  prop_results_text_reactive <- reactive({
    if(is.null(values$prop_test_results)) return("Belum ada hasil uji proporsi.")
    
    hasil <- values$prop_test_results
    p_val <- hasil$test$p.value
    
    test_output <- capture.output(print(hasil$test))
    
    interp_text <- switch(hasil$type,
                          "prop_one" = paste("H0: p =", input$prop_value, "\nH1: p ≠", input$prop_value),
                          "prop_two" = "H0: p1 = p2\nH1: p1 ≠ p2",
                          "chisq_gof" = "H0: Data mengikuti distribusi yang diharapkan\nH1: Data tidak mengikuti distribusi yang diharapkan",
                          "chisq_indep" = "H0: Variabel independen (tidak ada asosiasi)\nH1: Variabel tidak independen (ada asosiasi)"
    )
    
    keputusan <- ifelse(p_val > hasil$alpha, "Gagal menolak H0", "Tolak H0")
    kesimpulan <- ifelse(p_val > hasil$alpha, "Tidak ada bukti yang cukup untuk menolak hipotesis nol", "Ada bukti yang cukup untuk menolak hipotesis nol")
    
    paste(
      "📊 HASIL UJI PROPORSI\n=====================\n",
      paste("Jenis Uji:", str_to_upper(gsub("_", " ", hasil$type))),
      paste("Variabel:", hasil$variable),
      if(!is.null(hasil$group_var) && hasil$type %in% c("prop_two", "chisq_indep")) paste("Variabel Grouping:", hasil$group_var) else "",
      paste("Tingkat Signifikansi:", hasil$alpha),
      "\n",
      paste(test_output, collapse = "\n"),
      "\n\n📊 INTERPRETASI:",
      interp_text,
      paste("\nKeputusan:", keputusan),
      paste("📈 Kesimpulan:", kesimpulan),
      sep = "\n"
    )
  })
  
  output$prop_test_results <- renderText({
    prop_results_text_reactive()
  })
  
  output$contingency_table <- DT::renderDataTable({
    if(is.null(values$prop_test_results) || !values$prop_test_results$type %in% c("prop_two", "chisq_indep")) return(DT::datatable(data.frame(Message = "Tabel kontingensi tidak tersedia untuk jenis uji ini.")))
    cont_table <- table(sovi_data[[values$prop_test_results$variable]], sovi_data[[values$prop_test_results$group_var]])
    DT::datatable(as.data.frame.matrix(addmargins(cont_table)), options = list(pageLength = 10))
  })
  
  prop_test_plot_object <- reactive({
    if(is.null(values$prop_test_results)) return(NULL)
    
    # This reactive will hold the code to generate the plot
    # It doesn't return a ggplot object because mosaicplot is base R
    function() {
      if(values$prop_test_results$type %in% c("prop_two", "chisq_indep")) {
        mosaicplot(table(sovi_data[[values$prop_test_results$variable]], sovi_data[[values$prop_test_results$group_var]]), main = "Mosaic Plot", color = viridis(3), las=2)
      } else {
        df <- as.data.frame(table(sovi_data[[values$prop_test_results$variable]]))
        names(df) <- c("Category", "Frequency")
        ggplot(df, aes(x=Category, y=Frequency, fill=Category)) +
          geom_bar(stat="identity") +
          theme_minimal() +
          labs(title = paste("Bar Plot untuk", values$prop_test_results$variable))
      }
    }
  })
  
  output$prop_test_plot <- renderPlot({
    plot_func <- prop_test_plot_object()
    if(is.null(plot_func)) {
      plot.new(); text(0.5, 0.5, "Belum ada hasil uji proporsi.", cex = 1.2, col = "gray")
    } else {
      # Execute the plot function
      p <- plot_func()
      # If it's a ggplot object, print it
      if(inherits(p, "ggplot")) {
        print(p)
      }
    }
  })
  
  prop_interpretation_text_reactive <- reactive({
    if(is.null(values$prop_test_results)) return("Belum ada hasil uji proporsi.")
    
    hasil <- values$prop_test_results
    test_desc <- switch(hasil$type,
                        "prop_one" = "menguji apakah proporsi populasi sama dengan nilai tertentu",
                        "prop_two" = "menguji apakah proporsi dua grup berbeda",
                        "chisq_gof" = "menguji apakah data mengikuti distribusi yang diharapkan",
                        "chisq_indep" = "menguji independensi antara dua variabel kategorikal"
    )
    
    effect_size_text <- ""
    if(hasil$type == "chisq_indep") {
      chi_sq <- hasil$test$statistic
      n <- sum(hasil$test$observed)
      min_dim <- min(dim(hasil$test$observed)) - 1
      if (n > 0 && min_dim > 0) {
        cramers_v <- sqrt(chi_sq / (n * min_dim))
        effect_interpretation <- if(cramers_v < 0.1) "kecil" else if(cramers_v < 0.3) "sedang" else if(cramers_v < 0.5) "besar" else "sangat besar"
        effect_size_text <- paste0(
          "\n\n📏 UKURAN EFEK (CRAMER'S V):\n",
          "Cramer's V = ", round(cramers_v, 3), "\n",
          "Interpretasi: Asosiasi ", effect_interpretation
        )
      }
    }
    
    paste(
      "📝 INTERPRETASI UJI PROPORSI\n",
      "============================\n\n",
      "🎯 Tujuan Uji:\n",
      "Uji ini ", test_desc, ".\n\n",
      "📊 Hasil Statistik:\n",
      "• Statistik uji = ", round(hasil$test$statistic, 4), "\n",
      "• df = ", ifelse(is.null(hasil$test$parameter), "N/A", hasil$test$parameter), "\n",
      "• p-value = ", hasil$test$p.value, "\n",
      "• α = ", hasil$alpha, "\n\n",
      "🔍 Interpretasi Praktis:\n",
      ifelse(hasil$test$p.value > hasil$alpha,
             "Tidak ada perbedaan/asosiasi yang signifikan secara statistik.",
             "Ada perbedaan/asosiasi yang signifikan secara statistik."),
      effect_size_text,
      "\n\n💡 Rekomendasi:\n",
      "• Periksa asumsi uji (frekuensi expected ≥ 5 untuk chi-square)\n",
      "• Pertimbangkan signifikansi praktis selain signifikansi statistik\n",
      "• Untuk chi-square, periksa residual untuk identifikasi sel yang berkontribusi\n",
      "• Laporkan ukuran efek untuk memberikan konteks hasil",
      sep=""
    )
  })
  
  output$prop_test_interpretation <- renderText({
    prop_interpretation_text_reactive()
  })
  
  # ===== UJI VARIANS =====
  
  observeEvent(input$run_var_test, {
    req(input$var_test_type, input$var_test_var)
    test_result <- NULL
    if(input$var_test_type == "var_one") {
      # Custom implementation for one-sample variance test
      var_data <- sovi_data[[input$var_test_var]]; var_data <- var_data[!is.na(var_data)]
      n <- length(var_data); sample_var <- var(var_data)
      chi_sq_stat <- (n - 1) * sample_var / input$sigma_squared
      p_value <- 2 * min(pchisq(chi_sq_stat, df = n-1), 1 - pchisq(chi_sq_stat, df = n-1))
      test_result <- list(statistic = chi_sq_stat, parameter = n-1, p.value = p_value, method = "One Sample Chi-square test for variance")
    } else if(input$var_test_type == "f_test") {
      req(input$var_group_var)
      group_var <- as.factor(sovi_data[[input$var_group_var]])
      test_result <- var.test(sovi_data[[input$var_test_var]] ~ group_var)
    } else { # Levene or Bartlett
      req(input$var_group_var)
      group_var <- as.factor(sovi_data[[input$var_group_var]])
      test_result <- switch(input$var_test_type,
                            "levene" = car::leveneTest(sovi_data[[input$var_test_var]], group_var),
                            "bartlett" = bartlett.test(sovi_data[[input$var_test_var]], group_var))
    }
    if(!is.null(test_result)) {
      values$var_test_results <- list(test = test_result, type = input$var_test_type, variable = input$var_test_var, group_var = input$var_group_var, alpha = input$var_alpha)
      showNotification("✅ Uji varians selesai!", type = "message")
    }
  })
  
  output$var_test_results <- renderPrint({
    if(is.null(values$var_test_results)) return("Belum ada hasil uji varians.")
    cat("📈 HASIL UJI VARIANS\n====================\n")
    print(values$var_test_results$test)
  })
  
  output$var_test_plot <- renderPlot({
    if(is.null(values$var_test_results) || values$var_test_results$type == "var_one") {
      plot.new(); text(0.5, 0.5, "Visualisasi hanya untuk uji 2+ sampel.", cex = 1.2, col = "gray")
      return()
    }
    group_var <- as.factor(sovi_data[[values$var_test_results$group_var]])
    boxplot(sovi_data[[values$var_test_results$variable]] ~ group_var, main = "Perbandingan Varians Grup", col = viridis(length(levels(group_var))))
  })
  
  output$var_test_interpretation <- renderText({
    if(is.null(values$var_test_results)) return("Belum ada hasil uji varians.")
    p_val <- if(values$var_test_results$type == "levene") values$var_test_results$test$`Pr(>F)`[1] else values$var_test_results$test$p.value
    paste0("📝 INTERPRETASI UJI VARIANS\n===========================\n\n🎯 Hasil Statistik:\n• p-value = ", round(p_val, 6), "\n• α = ", values$var_test_results$alpha, "\n\n🔍 Interpretasi Praktis:\n", ifelse(p_val > values$var_test_results$alpha, "Varians tidak berbeda signifikan (homogen).", "Varians berbeda signifikan (tidak homogen)."))
  })
  
  # ===== ANOVA =====
  
  observeEvent(input$run_anova, {
    req(input$anova_response, input$anova_factor1)
    response_var <- sovi_data[[input$anova_response]]
    factor1 <- as.factor(sovi_data[[input$anova_factor1]])
    anova_result <- NULL; posthoc_result <- NULL; assumptions_result <- NULL
    
    formula_str <- "response_var ~ factor1"
    if(input$anova_type == "two_way") {
      req(input$anova_factor2)
      factor2 <- as.factor(sovi_data[[input$anova_factor2]])
      formula_str <- paste(formula_str, if(input$anova_interaction) "*" else "+", "factor2")
    } else if(input$anova_type == "repeated") {
      req(input$anova_subject)
      subject_var <- as.factor(sovi_data[[input$anova_subject]])
      formula_str <- paste(formula_str, "+ Error(subject_var/factor1)")
    }
    
    anova_result <- aov(as.formula(formula_str))
    
    if(input$anova_posthoc && input$anova_type != "repeated") {
      posthoc_result <- switch(input$posthoc_method,
                               "tukey" = TukeyHSD(anova_result),
                               "bonferroni" = pairwise.t.test(response_var, factor1, p.adjust.method = "bonferroni"),
                               "scheffe" = pairwise.t.test(response_var, factor1, p.adjust.method = "none"), # Simplified
                               "lsd" = pairwise.t.test(response_var, factor1, p.adjust.method = "none"))
    }
    
    if(input$anova_assumptions) {
      residuals <- residuals(anova_result)
      norm_test <- if(length(residuals) <= 5000) shapiro.test(residuals) else lillie.test(residuals)
      homo_test <- car::leveneTest(response_var, factor1)
      assumptions_result <- list(normality = norm_test, homogeneity = homo_test)
    }
    
    values$anova_results <- list(model = anova_result, type = input$anova_type, response = input$anova_response, factor1 = input$anova_factor1, factor2 = input$anova_factor2, posthoc = posthoc_result, assumptions = assumptions_result, alpha = input$anova_alpha)
    showNotification("✅ ANOVA selesai!", type = "message")
  })
  
  output$anova_results <- renderPrint({
    if(is.null(values$anova_results)) return("Belum ada hasil ANOVA.")
    cat("📊 HASIL ANOVA\n==============\n")
    print(summary(values$anova_results$model))
  })
  
  output$anova_plot <- renderPlot({
    if(is.null(values$anova_results)) {
      plot.new(); text(0.5, 0.5, "Belum ada hasil ANOVA.", cex = 1.2, col = "gray")
      return()
    }
    if(values$anova_results$type == "two_way") {
      interaction.plot(as.factor(sovi_data[[values$anova_results$factor1]]), as.factor(sovi_data[[values$anova_results$factor2]]), sovi_data[[values$anova_results$response]], main = "Interaction Plot", col = viridis(3), trace.label = values$anova_results$factor2)
    } else {
      par(mfrow=c(1,2))
      boxplot(sovi_data[[values$anova_results$response]] ~ as.factor(sovi_data[[values$anova_results$factor1]]), main = "Boxplot per Grup", col = viridis(length(unique(sovi_data[[values$anova_results$factor1]]))))
      plot(values$anova_results$model, 2)
    }
  })
  
  output$posthoc_results <- renderPrint({
    if(is.null(values$anova_results) || is.null(values$anova_results$posthoc)) return("Tidak ada hasil post-hoc.")
    cat("🔍 HASIL UJI POST-HOC\n=====================\n")
    print(values$anova_results$posthoc)
  })
  
  output$posthoc_plot <- renderPlot({
    if(is.null(values$anova_results) || is.null(values$anova_results$posthoc) || input$posthoc_method != "tukey") {
      plot.new(); text(0.5, 0.5, "Plot hanya untuk Tukey HSD.", cex = 1.2, col = "gray")
      return()
    }
    plot(values$anova_results$posthoc, las=1)
  })
  
  output$anova_assumptions_results <- renderPrint({
    if(is.null(values$anova_results) || is.null(values$anova_results$assumptions)) return("Tidak ada hasil uji asumsi.")
    cat("✅ HASIL UJI ASUMSI ANOVA\n=========================\n\n1. UJI NORMALITAS RESIDUAL:\n")
    print(values$anova_results$assumptions$normality)
    cat("\n2. UJI HOMOGENITAS VARIANS:\n")
    print(values$anova_results$assumptions$homogeneity)
  })
  
  output$anova_assumptions_plots <- renderPlot({
    if(is.null(values$anova_results) || is.null(values$anova_results$assumptions)) {
      plot.new(); text(0.5, 0.5, "Tidak ada plot asumsi.", cex = 1.2, col = "gray")
      return()
    }
    par(mfrow = c(1, 2))
    plot(values$anova_results$model, 1) # Residuals vs Fitted
    plot(values$anova_results$model, 2) # Q-Q Plot
  })
  
  output$anova_interpretation <- renderText({
    if(is.null(values$anova_results)) return("Belum ada hasil ANOVA.")
    p_val <- summary(values$anova_results$model)[[1]]$`Pr(>F)`[1]
    paste0("📝 INTERPRETASI ANOVA\n=====================\n\n🎯 Hasil Statistik:\n• p-value = ", round(p_val, 6), "\n• α = ", values$anova_results$alpha, "\n\n🔍 Interpretasi Praktis:\n", ifelse(p_val > values$anova_results$alpha, "Tidak ada perbedaan signifikan antar grup.", "Ada perbedaan signifikan antar grup. Lanjutkan dengan uji post-hoc."))
  })
  
  # ===== REGRESI =====
  
  observeEvent(input$run_regression, {
    req(input$reg_response, input$reg_predictors)
    reg_data <- sovi_data[, c(input$reg_response, input$reg_predictors), drop = FALSE]
    reg_data <- reg_data[complete.cases(reg_data), ]
    
    formula_str <- paste(input$reg_response, "~", paste(input$reg_predictors, collapse = " + "))
    if(input$reg_interaction && length(input$reg_predictors) >= 2) formula_str <- paste(input$reg_response, "~", paste(input$reg_predictors, collapse = " * "))
    if(!input$reg_intercept) formula_str <- paste(formula_str, "- 1")
    formula_obj <- as.formula(formula_str)
    
    full_model <- lm(formula_obj, data = reg_data)
    reg_model <- switch(input$reg_method,
                        "enter" = full_model,
                        "forward" = step(lm(paste(input$reg_response, "~ 1"), data=reg_data), scope=list(lower=~1, upper=formula(full_model)), direction="forward", trace=0),
                        "backward" = step(full_model, direction="backward", trace=0),
                        "stepwise" = step(lm(paste(input$reg_response, "~ 1"), data=reg_data), scope=list(lower=~1, upper=formula(full_model)), direction="both", trace=0))
    
    diagnostics <- NULL
    if(input$reg_diagnostics) {
      residuals_data <- residuals(reg_model)
      norm_test <- if(length(residuals_data) <= 5000) shapiro.test(residuals_data) else lillie.test(residuals_data)
      bp_test <- lmtest::bptest(reg_model)
      dw_test <- lmtest::dwtest(reg_model)
      vif_values <- if(length(coef(reg_model)) > 2) car::vif(reg_model) else NA
      diagnostics <- list(normality = norm_test, heteroscedasticity = bp_test, autocorrelation = dw_test, multicollinearity = vif_values)
    }
    
    values$regression_model <- list(model = reg_model, data = reg_data, formula = formula_str, method = input$reg_method, diagnostics = diagnostics, standardized = input$reg_standardize)
    showNotification("✅ Regresi linear berganda selesai!", type = "message")
  })
  
  output$regression_summary <- renderPrint({
    if(is.null(values$regression_model)) return("Belum ada hasil regresi.")
    cat("📈 HASIL REGRESI LINEAR BERGANDA\n================================\n")
    print(summary(values$regression_model$model))
  })
  
  output$regression_diagnostics <- renderPlot({
    if(is.null(values$regression_model)) {
      plot.new(); text(0.5, 0.5, "Belum ada hasil regresi.", cex = 1.2, col = "gray")
      return()
    }
    par(mfrow = c(2, 2)); plot(values$regression_model$model)
  })
  
  output$regression_assumptions <- renderPrint({
    if(is.null(values$regression_model) || is.null(values$regression_model$diagnostics)) return("Tidak ada hasil uji asumsi.")
    diag <- values$regression_model$diagnostics
    cat("✅ HASIL UJI ASUMSI REGRESI\n===========================\n\n1. Normalitas Residual:\n"); print(diag$normality)
    cat("\n2. Heteroskedastisitas (Breusch-Pagan):\n"); print(diag$heteroscedasticity)
    cat("\n3. Autokorelasi (Durbin-Watson):\n"); print(diag$autocorrelation)
    cat("\n4. Multikolinearitas (VIF):\n"); print(diag$multicollinearity)
  })
  
  output$regression_assumptions_plots <- renderPlot({
    if(is.null(values$regression_model)) {
      plot.new(); text(0.5, 0.5, "Belum ada hasil regresi.", cex = 1.2, col = "gray")
      return()
    }
    par(mfrow = c(2, 2)); plot(values$regression_model$model, which = 1:4)
  })
  
  output$model_comparison <- DT::renderDataTable({
    if(is.null(values$regression_model)) return(DT::datatable(data.frame(Message = "Belum ada model untuk dibandingkan.")))
    model <- values$regression_model$model
    simple_model <- lm(paste(input$reg_response, "~", input$reg_predictors[1]), data = values$regression_model$data)
    comparison_df <- data.frame(
      Model = c("Simple", "Full (Selected)"),
      R_squared = c(summary(simple_model)$r.squared, summary(model)$r.squared),
      Adj_R_squared = c(summary(simple_model)$adj.r.squared, summary(model)$adj.r.squared),
      AIC = c(AIC(simple_model), AIC(model)),
      BIC = c(BIC(simple_model), BIC(model))
    )
    DT::datatable(comparison_df, options = list(pageLength = 10)) %>% DT::formatRound(columns = 2:5, digits = 4)
  })
  
  output$model_comparison_plot <- renderPlot({
    if(is.null(values$regression_model)) {
      plot.new(); text(0.5, 0.5, "Belum ada model untuk dibandingkan.", cex = 1.2, col = "gray")
      return()
    }
    model <- values$regression_model$model
    actual <- values$regression_model$data[[input$reg_response]]
    predicted <- fitted(model)
    plot(actual, predicted, main = "Actual vs Predicted", xlab = "Actual", ylab = "Predicted", pch = 16, col = "#5e72e4")
    abline(0, 1, col = "#f5365c", lwd = 2)
  })
  
  output$prediction_inputs <- renderUI({
    if(is.null(values$regression_model)) return(p("Belum ada model untuk prediksi."))
    predictors <- names(coef(values$regression_model$model))[-1] # Get predictors from the final model
    lapply(predictors, function(p) {
      numericInput(paste0("pred_", p), label = p, value = mean(sovi_data[[p]], na.rm = TRUE))
    })
  })
  
  prediction_result <- eventReactive(input$make_prediction, {
    if(is.null(values$regression_model)) return(NULL)
    predictors <- names(coef(values$regression_model$model))[-1]
    pred_data <- data.frame(lapply(predictors, function(p) input[[paste0("pred_", p)]]))
    names(pred_data) <- predictors
    predict(values$regression_model$model, newdata = pred_data, interval = "prediction")
  })
  
  output$prediction_results <- renderPrint({
    pred <- prediction_result()
    if(is.null(pred)) return("Masukkan nilai prediktor dan klik 'Buat Prediksi'.")
    cat("🔮 HASIL PREDIKSI\n=================\n\nPrediksi:", round(pred[1], 4), "\n95% Prediction Interval: [", round(pred[2], 4), ",", round(pred[3], 4), "]\n")
  })
  
  output$prediction_plot <- renderPlot({
    if(is.null(values$regression_model)) {
      plot.new(); text(0.5, 0.5, "Belum ada model untuk prediksi.", cex = 1.2, col = "gray")
      return()
    }
    # Simplified plot
    plot(values$regression_model$data[[input$reg_response]], fitted(values$regression_model$model), main="Actual vs. Fitted", xlab="Actual", ylab="Fitted")
    abline(0,1, col="red")
  })
  
  output$regression_interpretation <- renderText({
    if(is.null(values$regression_model)) return("Belum ada hasil regresi.")
    model_summary <- summary(values$regression_model$model)
    paste0(
      "📝 INTERPRETASI REGRESI\n======================\n\n📊 Kualitas Model:\n• R² = ", round(model_summary$r.squared, 4), " (", round(model_summary$r.squared*100, 1), "% varians dijelaskan)\n• p-value (model) = ", format(pf(model_summary$fstatistic[1], model_summary$fstatistic[2], model_summary$fstatistic[3], lower.tail = FALSE), scientific = TRUE), "\n\n🔍 Interpretasi Koefisien:\n• p-value < 0.05 menunjukkan prediktor signifikan.\n• Koefisien positif/negatif menunjukkan arah hubungan.\n\n💡 Rekomendasi:\n• Periksa asumsi regresi (normalitas, homoskedastisitas, dll).\n• Evaluasi multikolinearitas (VIF).\n• Validasi model dengan data baru jika memungkinkan."
    )
  })
  
  # ===== DOWNLOAD HANDLERS UNTUK BERANDA =====
  
  # Helper function untuk membuat dokumen Word
  create_word_doc <- function(title, content) {
    doc <- officer::read_docx()
    doc <- doc %>%
      officer::body_add_par(title, style = "heading 1") %>%
      officer::body_add_par("") %>%
      officer::body_add_par(content, style = "Normal")
    return(doc)
  }
  
  output$download_intro_card <- downloadHandler(
    filename = function() {
      paste0("Intro_Dashboard_SAVI_", Sys.Date(), ".docx")
    },
    content = function(file) {
      content <- paste(
        "Dashboard SAVI - Sistem Analisis Visualisasi Interaktif\n",
        "========================================================\n\n",
        "Dashboard ini dibuat untuk memenuhi persyaratan UAS Komputasi Statistik STIS - Semester Genap TA. 2024/2025\n\n",
        "METADATA DATASET:\n",
        "-----------------\n",
        "• Dataset SoVI: https://raw.githubusercontent.com/bmlmcmc/naspaclust/main/data/sovi_data.csv\n",
        "• Matriks Jarak: https://raw.githubusercontent.com/bmlmcmc/naspaclust/main/data/distance.csv\n",
        "• Publikasi: https://www.sciencedirect.com/science/article/pii/S2352340921010180\n\n",
        "TUJUAN DASHBOARD:\n",
        "-----------------\n",
        "• Menganalisis data Indeks Kerentanan Sosial (SoVI) secara komprehensif\n",
        "• Menerapkan berbagai metode statistik inferensia dan deskriptif\n",
        "• Melakukan visualisasi data interaktif\n",
        "• Mengimplementasikan uji asumsi statistik yang diperlukan\n",
        "• Menyediakan platform analisis yang user-friendly untuk pembelajaran\n"
      )
      doc <- create_word_doc("Informasi Pengantar Dashboard SAVI", content)
      print(doc, target = file)
    }
  )
  
  output$download_sovi_summary <- downloadHandler(
    filename = function() {
      paste0("Ringkasan_Data_SoVI_", Sys.Date(), ".docx")
    },
    content = function(file) {
      req(sovi_data)
      
      summary_df <- data.frame(
        Variabel = names(sovi_data),
        Tipe = sapply(sovi_data, class),
        Missing = sapply(sovi_data, function(x) sum(is.na(x))),
        Min = sapply(sovi_data, function(x) if(is.numeric(x)) round(min(x, na.rm = TRUE), 3) else "-"),
        Max = sapply(sovi_data, function(x) if(is.numeric(x)) round(max(x, na.rm = TRUE), 3) else "-"),
        Mean = sapply(sovi_data, function(x) if(is.numeric(x)) round(mean(x, na.rm = TRUE), 3) else "-"),
        stringsAsFactors = FALSE
      )
      
      ft <- flextable(summary_df) %>% autofit() %>% theme_booktabs()
      
      doc <- read_docx() %>%
        body_add_par("Ringkasan Data SoVI", style = "heading 1") %>%
        body_add_par(paste("Tanggal Dibuat:", Sys.Date())) %>%
        body_add_par("") %>%
        body_add_flextable(ft)
      
      print(doc, target = file)
    }
  )
  
  output$download_quick_viz <- downloadHandler(
    filename = function() {
      paste0("Visualisasi_Cepat_", Sys.Date(), ".jpg")
    },
    content = function(file) {
      req(sovi_data)
      numeric_vars <- names(sovi_data)[sapply(sovi_data, is.numeric)]
      if (length(numeric_vars) > 0) {
        var_name <- numeric_vars[1]
        p <- ggplot(sovi_data, aes_string(x = var_name)) +
          geom_histogram(fill = "#5e72e4", alpha = 0.7, bins = 30) +
          theme_minimal() +
          labs(title = paste("Distribusi", var_name), x = var_name, y = "Frekuensi") +
          theme(plot.title = element_text(hjust = 0.5))
        ggsave(file, plot = p, device = "jpeg", width = 8, height = 6, dpi = 300)
      }
    }
  )
  
  output$download_distance_info <- downloadHandler(
    filename = function() {
      paste0("Info_Matriks_Jarak_", Sys.Date(), ".docx")
    },
    content = function(file) {
      req(distance_matrix)
      content <- paste(
        "INFORMASI MATRIKS JARAK\n",
        "========================\n\n",
        paste("Dimensi Matriks:", nrow(distance_matrix), "x", ncol(distance_matrix)),
        paste("Tipe Data: Matriks Penimbang Jarak"),
        "\nStatistik Deskriptif Jarak:\n",
        paste("• Jarak Minimum:", round(min(distance_matrix, na.rm = TRUE), 3)),
        paste("• Jarak Maksimum:", round(max(distance_matrix, na.rm = TRUE), 3)),
        paste("• Jarak Rata-rata:", round(mean(as.matrix(distance_matrix), na.rm = TRUE), 3)),
        sep = "\n"
      )
      doc <- create_word_doc("Informasi Matriks Jarak", content)
      print(doc, target = file)
    }
  )
  
  output$download_distance_stats <- downloadHandler(
    filename = function() {
      paste0("Statistik_Jarak_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      req(distance_matrix)
      stats_data <- data.frame(
        ID = 1:nrow(distance_matrix),
        Min_Distance = apply(distance_matrix, 1, function(x) round(min(x[x > 0], na.rm = TRUE), 3)),
        Max_Distance = apply(distance_matrix, 1, function(x) round(max(x, na.rm = TRUE), 3)),
        Mean_Distance = apply(distance_matrix, 1, function(x) round(mean(x, na.rm = TRUE), 3)),
        SD_Distance = apply(distance_matrix, 1, function(x) round(sd(x, na.rm = TRUE), 3))
      )
      write.xlsx(stats_data, file)
    }
  )
  
  output$download_usage_info <- downloadHandler(
    filename = function() {
      paste0("Info_Penggunaan_Matriks_Jarak_", Sys.Date(), ".docx")
    },
    content = function(file) {
      content <- paste(
        "APLIKASI MATRIKS JARAK DALAM ANALISIS DATA\n",
        "===========================================\n\n",
        "Matriks jarak adalah alat fundamental dalam berbagai teknik analisis data, terutama untuk data multivariat dan spasial. Berikut adalah beberapa aplikasinya:\n\n",
        "1. Analisis Clustering:\n",
        "   - Mengelompokkan observasi yang 'dekat' atau serupa berdasarkan jarak.\n",
        "   - Metode seperti Hierarchical Clustering dan K-Medoids (PAM) secara langsung menggunakan matriks jarak.\n\n",
        "2. Analisis Spasial:\n",
        "   - Menganalisis pola geografis dan hubungan spasial antar wilayah.\n",
        "   - Digunakan dalam perhitungan Indeks Moran untuk autokorelasi spasial.\n\n",
        "3. Multidimensional Scaling (MDS):\n",
        "   - Merepresentasikan jarak antar item dalam ruang dimensi yang lebih rendah (biasanya 2D atau 3D) untuk visualisasi.\n",
        "   - Membantu memahami struktur tersembunyi dalam data.\n\n",
        "4. Deteksi Outlier:\n",
        "   - Observasi yang memiliki jarak rata-rata yang sangat besar ke semua observasi lain dapat diidentifikasi sebagai outlier.\n\n",
        "5. Analisis Kemiripan:\n",
        "   - Mengukur secara kuantitatif tingkat kemiripan atau perbedaan antar unit observasi."
      )
      doc <- create_word_doc("Informasi Penggunaan Matriks Jarak", content)
      print(doc, target = file)
    }
  )
  
  output$download_metadata_complete <- downloadHandler(
    filename = function() {
      paste0("Metadata_Lengkap_", Sys.Date(), ".docx")
    },
    content = function(file) {
      content <- paste(
        "METADATA LENGKAP DASHBOARD SAVI\n",
        "================================\n\n",
        "INFORMASI MATA KULIAH:\n",
        "----------------------\n",
        "• Mata Kuliah: Komputasi Statistik (K201313/3 SKS)\n",
        "• Dosen: Robert Kurniawan, Sukim, Yuliagnis Transver Wijaya\n",
        "• UAS: Rabu, 23 Juli 2025, 10.30-12.30 WIB\n",
        "• Institusi: Sekolah Tinggi Ilmu Statistik (STIS)\n\n",
        "SUMBER DATA:\n",
        "------------\n",
        "• Dataset SoVI: https://raw.githubusercontent.com/bmlmcmc/naspaclust/main/data/sovi_data.csv\n",
        "• Matriks Jarak: https://raw.githubusercontent.com/bmlmcmc/naspaclust/main/data/distance.csv\n",
        "• Referensi: https://www.sciencedirect.com/science/article/pii/S2352340921010180"
      )
      doc <- create_word_doc("Metadata Lengkap Dashboard SAVI", content)
      print(doc, target = file)
    }
  )
  
  output$download_sovi_data <- downloadHandler(
    filename = function() {
      paste0("Data_SoVI_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      write.xlsx(sovi_data, file)
    }
  )
  
  output$download_distance_matrix <- downloadHandler(
    filename = function() {
      paste0("Matriks_Jarak_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      write.xlsx(distance_matrix, file, rowNames = TRUE)
    }
  )
  
  output$download_combined_data <- downloadHandler(
    filename = function() {
      paste0("Data_Gabungan_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      wb <- createWorkbook()
      addWorksheet(wb, "Data_SoVI")
      writeData(wb, "Data_SoVI", sovi_data)
      addWorksheet(wb, "Matriks_Jarak")
      writeData(wb, "Matriks_Jarak", distance_matrix, rowNames = TRUE)
      saveWorkbook(wb, file, overwrite = TRUE)
    }
  )
  
  # ===== DOWNLOAD HANDLERS UNTUK MANAJEMEN DATA =====
  
  output$download_categorized_data <- downloadHandler(
    filename = function() {
      paste0("Data_Terkategorisasi_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      if(is.null(values$categorized_data)) {
        df <- data.frame(Pesan = "Tidak ada data untuk diunduh. Silakan terapkan kategorisasi terlebih dahulu.")
        write.xlsx(df, file)
      } else {
        write.xlsx(values$categorized_data, file)
      }
    }
  )
  
  output$download_cat_interpretation <- downloadHandler(
    filename = function() {
      paste0("Interpretasi_Kategorisasi_", Sys.Date(), ".docx")
    },
    content = function(file) {
      # Menggunakan reactive expression untuk mendapatkan konten
      content <- cat_interpretation_reactive()
      doc <- create_word_doc("Interpretasi Kategorisasi Data", content)
      print(doc, target = file)
    }
  )
  
  output$download_cat_report <- downloadHandler(
    filename = function() {
      paste0("Laporan_Kategorisasi_", Sys.Date(), ".docx")
    },
    content = function(file) {
      req(values$categorized_data)
      
      doc <- read_docx() %>%
        body_add_par("Laporan Lengkap Kategorisasi Data", style = "heading 1") %>%
        body_add_par(paste("Tanggal:", Sys.Date())) %>%
        body_add_par("") %>%
        body_add_par("Ringkasan", style = "heading 2") %>%
        body_add_par(paste0("Variabel '", input$cat_variable, "' telah dikategorikan menjadi ", input$cat_groups, " grup menggunakan metode '", input$cat_method, "'.")) %>%
        body_add_par("") %>%
        body_add_par("Statistik per Kategori", style = "heading 2")
      
      # Ambil data statistik
      cat_var <- paste0(input$cat_variable, "_categorized")
      cat_stats <- values$categorized_data %>%
        group_by_at(cat_var) %>%
        summarise(
          N = n(),
          Mean_Original = round(mean(get(input$cat_variable), na.rm = TRUE), 4),
          SD_Original = round(sd(get(input$cat_variable), na.rm = TRUE), 4),
          .groups = 'drop'
        )
      
      # Tambahkan tabel ke dokumen
      ft <- flextable(cat_stats) %>% autofit() %>% theme_booktabs()
      doc <- doc %>% body_add_flextable(ft) %>% body_add_par("")
      
      # Tambahkan plot
      temp_plot_file <- tempfile(fileext = ".png")
      
      # Membuat plot object dari reactive
      p_combined <- categorization_plot_object()
      
      # Menyimpan plot ke file sementara
      # Perlu menggunakan grid::grid.draw() untuk menyimpan objek grid
      png(temp_plot_file, width = 10, height = 5, units = "in", res = 300)
      grid::grid.draw(p_combined)
      dev.off()
      
      doc <- doc %>%
        body_add_par("Visualisasi Perbandingan", style = "heading 2") %>%
        body_add_img(src = temp_plot_file, width = 6, height = 3)
      
      print(doc, target = file)
    }
  )
  
  output$download_cat_plot_jpg <- downloadHandler(
    filename = function() {
      paste0("Plot_Kategorisasi_", Sys.Date(), ".jpg")
    },
    content = function(file) {
      req(values$categorized_data)
      p_combined <- categorization_plot_object()
      # Menggunakan ggsave dengan grid object memerlukan cara berbeda
      ggsave(file, plot = p_combined, device = "jpeg", width = 10, height = 5, dpi = 300)
    }
  )
  
  output$download_cat_stats <- downloadHandler(
    filename = function() {
      paste0("Statistik_Kategori_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      req(values$categorized_data)
      
      cat_var <- paste0(input$cat_variable, "_categorized")
      cat_stats <- values$categorized_data %>%
        group_by_at(cat_var) %>%
        summarise(
          N = n(),
          Mean_Original = round(mean(get(input$cat_variable), na.rm = TRUE), 4),
          Median_Original = round(median(get(input$cat_variable), na.rm = TRUE), 4),
          SD_Original = round(sd(get(input$cat_variable), na.rm = TRUE), 4),
          Min_Original = round(min(get(input$cat_variable), na.rm = TRUE), 4),
          Max_Original = round(max(get(input$cat_variable), na.rm = TRUE), 4),
          .groups = 'drop'
        )
      write.xlsx(cat_stats, file)
    }
  )
  
  # ===== DOWNLOAD HANDLERS UNTUK STATISTIK DESKRIPTIF =====
  
  output$download_desc_excel <- downloadHandler(
    filename = function() paste0("Statistik_Deskriptif_", Sys.Date(), ".xlsx"),
    content = function(file) {
      req(values$descriptive_results)
      write.xlsx(values$descriptive_results, file)
    }
  )
  
  output$download_desc_interpretation <- downloadHandler(
    filename = function() paste0("Interpretasi_Deskriptif_", Sys.Date(), ".docx"),
    content = function(file) {
      content <- desc_interpretation_reactive()
      doc <- create_word_doc("Interpretasi Statistik Deskriptif", content)
      print(doc, target = file)
    }
  )
  
  output$download_desc_report <- downloadHandler(
    filename = function() paste0("Laporan_Deskriptif_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$descriptive_results)
      
      doc <- read_docx() %>%
        body_add_par("Laporan Lengkap Statistik Deskriptif", style = "heading 1") %>%
        body_add_par(paste("Tanggal:", Sys.Date())) %>%
        body_add_par("") %>%
        body_add_par("Tabel Statistik", style = "heading 2")
      
      ft <- flextable(values$descriptive_results) %>% autofit() %>% theme_booktabs()
      doc <- doc %>% body_add_flextable(ft) %>% body_add_par("")
      
      doc <- doc %>% 
        body_add_par("Interpretasi", style = "heading 2") %>%
        body_add_par(desc_interpretation_reactive())
      
      print(doc, target = file)
    }
  )
  
  output$download_desc_plots_jpg <- downloadHandler(
    filename = function() paste0("Plot_Distribusi_", Sys.Date(), ".jpg"),
    content = function(file) {
      req(descriptive_plots_object())
      p <- descriptive_plots_object()
      ggsave(file, plot = p, device = "jpeg", width = 10, height = 8, dpi = 300)
    }
  )
  
  output$download_group_comp_plot_jpg <- downloadHandler(
    filename = function() paste0("Plot_Perbandingan_Grup_", Sys.Date(), ".jpg"),
    content = function(file) {
      req(group_comparison_plot_object())
      p <- group_comparison_plot_object()
      ggsave(file, plot = p, device = "jpeg", width = 8, height = 6, dpi = 300)
    }
  )
  
  # ===== DOWNLOAD HANDLERS UNTUK ANALISIS MATRIKS JARAK =====
  
  output$download_distance_results <- downloadHandler(
    filename = function() {
      paste0("Hasil_Analisis_Jarak_", Sys.Date(), ".xlsx")
    },
    content = function(file) {
      req(values$distance_analysis_results)
      result <- values$distance_analysis_results
      
      # Menyiapkan data untuk diunduh berdasarkan jenis analisis
      data_to_download <- switch(result$type,
                                 "descriptive" = result$stats,
                                 "clustering" = {
                                   # Menggabungkan hasil cluster dengan data asli
                                   data_with_clusters <- sovi_data
                                   data_with_clusters$Cluster <- result$clusters
                                   data_with_clusters
                                 },
                                 "pca" = {
                                   # Menggabungkan koordinat MDS/PCA dengan data asli
                                   data_with_coords <- sovi_data
                                   coords_df <- as.data.frame(result$coordinates)
                                   names(coords_df) <- paste0("Dim", 1:ncol(coords_df))
                                   cbind(data_with_coords, coords_df)
                                 },
                                 # Default case jika tidak ada data tabel
                                 data.frame(Message = "Tidak ada data tabel untuk jenis analisis ini.")
      )
      
      # Menulis ke file Excel
      write.xlsx(data_to_download, file, rowNames = FALSE)
    }
  )
  
  output$download_distance_plot <- downloadHandler(
    filename = function() {
      paste0("Plot_Analisis_Jarak_", Sys.Date(), ".jpg")
    },
    content = function(file) {
      req(values$distance_analysis_results)
      result <- values$distance_analysis_results
      
      # Menggunakan ggsave untuk plot ggplot atau jpeg untuk base plot
      if (result$type %in% c("clustering", "pca")) {
        # Membuat ulang plot ggplot
        p <- if (result$type == "clustering") {
          mds_coords <- cmdscale(as.matrix(distance_matrix), k = 2)
          plot_data <- data.frame(X = mds_coords[,1], Y = mds_coords[,2], Cluster = as.factor(result$clusters))
          ggplot(plot_data, aes(x = X, y = Y, color = Cluster)) +
            geom_point(size = 3, alpha = 0.7) +
            scale_color_viridis_d() +
            theme_minimal() +
            labs(title = paste("Hasil Clustering -", str_to_title(result$method)))
        } else { # PCA
          plot_data <- data.frame(Component = 1:length(result$eigenvalues), Eigenvalue = result$eigenvalues)
          ggplot(plot_data, aes(x = Component, y = Eigenvalue)) +
            geom_line(color = "#5e72e4", size = 1) + geom_point(color = "#11cdef", size = 3) +
            theme_minimal() + labs(title = "Scree Plot - Analisis Komponen Utama")
        }
        ggsave(file, plot = p, device = "jpeg", width = 8, height = 6, dpi = 300)
        
      } else {
        # Menggunakan jpeg() untuk base R plot (histogram atau heatmap)
        jpeg(file, width = 800, height = 600, quality = 95)
        if (result$type == "descriptive") {
          hist(result$values, breaks = 50, col = "#5e72e4", main = "Distribusi Jarak Antar Observasi", xlab = "Jarak")
        } else if (result$type == "heatmap") {
          # Mengambil sampel jika matriks terlalu besar untuk visualisasi
          distance_matrix_num <- as.matrix(distance_matrix)
          if(nrow(distance_matrix_num) > 100) {
            indices <- sample(1:nrow(distance_matrix_num), 100)
            distance_matrix_num <- distance_matrix_num[indices, indices]
          }
          heatmap(distance_matrix_num, symm = TRUE, main = "Heatmap Matriks Jarak (Sampel)", col = viridis(256))
        }
        dev.off()
      }
    }
  )
  
  output$download_distance_interpretation <- downloadHandler(
    filename = function() {
      paste0("Interpretasi_Analisis_Jarak_", Sys.Date(), ".docx")
    },
    content = function(file) {
      req(values$distance_analysis_results)
      result <- values$distance_analysis_results
      
      # Menghasilkan teks interpretasi berdasarkan jenis analisis
      interpretation_text <- switch(result$type,
                                    "descriptive" = paste(
                                      "INTERPRETASI STATISTIK DESKRIPTIF MATRIKS JARAK\n\n",
                                      "Analisis ini memberikan ringkasan dari semua jarak antar observasi dalam dataset.\n",
                                      "- Mean (Rata-rata): Jarak rata-rata antar semua pasangan titik. Nilai yang lebih tinggi menunjukkan data secara keseluruhan lebih tersebar (heterogen).\n",
                                      "- SD (Standar Deviasi): Seberapa bervariasi jarak-jarak tersebut. SD yang besar menunjukkan beberapa titik sangat jauh sementara yang lain sangat dekat.\n",
                                      "- Min/Max: Jarak terpendek dan terjauh yang ditemukan antara dua titik mana pun."
                                    ),
                                    "clustering" = paste(
                                      "INTERPRETASI HASIL CLUSTERING\n\n",
                                      "Metode:", str_to_title(result$method), "\n",
                                      "Jumlah Cluster:", result$n_clusters, "\n\n",
                                      "Analisis ini mengelompokkan observasi ke dalam", result$n_clusters, "grup berdasarkan kedekatan jaraknya. Observasi dalam satu cluster memiliki jarak yang lebih dekat satu sama lain dibandingkan dengan observasi di cluster lain.\n\n",
                                      "Visualisasi plot menunjukkan representasi 2D dari jarak antar titik (menggunakan MDS). Titik-titik dengan warna yang sama adalah anggota dari cluster yang sama. Idealnya, cluster akan terlihat sebagai gumpalan warna yang terpisah."
                                    ),
                                    "pca" = paste(
                                      "INTERPRETASI ANALISIS KOMPONEN UTAMA (PCA/MDS)\n\n",
                                      "Analisis ini (sering disebut Multidimensional Scaling saat diterapkan pada matriks jarak) bertujuan untuk mereduksi dimensi data sambil mempertahankan informasi jarak sebanyak mungkin.\n\n",
                                      "- Eigenvalue: Mewakili jumlah varians yang dijelaskan oleh setiap komponen (dimensi). Komponen dengan eigenvalue tertinggi adalah yang paling penting.\n",
                                      "- Scree Plot: Visualisasi dari eigenvalue. Titik 'patah' (elbow) pada plot dapat membantu menentukan jumlah komponen yang signifikan untuk dipertahankan."
                                    ),
                                    "heatmap" = paste(
                                      "INTERPRETASI HEATMAP MATRIKS JARAK\n\n",
                                      "Heatmap ini memvisualisasikan seluruh matriks jarak. Setiap sel mewakili jarak antara dua observasi.\n",
                                      "- Warna: Skema warna digunakan untuk menunjukkan jarak. Warna yang lebih 'dingin' (misal: biru/ungu) biasanya menunjukkan jarak yang dekat, sementara warna yang lebih 'panas' (misal: kuning/merah) menunjukkan jarak yang jauh.\n",
                                      "- Pola: Adanya blok-blok berwarna solid menunjukkan adanya kelompok (cluster) observasi yang saling berdekatan."
                                    )
      )
      
      # Membuat dokumen Word
      doc <- create_word_doc("Interpretasi Analisis Matriks Jarak", interpretation_text)
      print(doc, target = file)
    }
  )
  
  # ===== DOWNLOAD HANDLERS UNTUK UJI ASUMSI =====
  
  output$download_norm_results <- downloadHandler(
    filename = function() paste0("Hasil_Uji_Normalitas_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$normality_results)
      
      doc <- read_docx() %>%
        body_add_par("Hasil Uji Normalitas", style = "heading 1") %>%
        body_add_par(paste("Variabel:", values$normality_results$variable)) %>%
        body_add_par("")
      
      for(test_name in names(values$normality_results$tests)) {
        test_result <- values$normality_results$tests[[test_name]]
        doc <- doc %>%
          body_add_par(paste(str_to_upper(test_name), "TEST"), style = "heading 2") %>%
          body_add_par(paste("p-value =", format(test_result$p.value, scientific = TRUE))) %>%
          body_add_par(paste("Kesimpulan:", ifelse(test_result$p.value > values$normality_results$alpha, "Normal", "Tidak Normal"))) %>%
          body_add_par("")
      }
      print(doc, target = file)
    }
  )
  
  output$download_norm_plots <- downloadHandler(
    filename = function() paste0("Plot_Normalitas_", Sys.Date(), ".jpg"),
    content = function(file) {
      req(values$normality_results)
      jpeg(file, width = 800, height = 400, quality = 95)
      par(mfrow = c(1, 2))
      var_data <- values$normality_results$data
      hist(var_data, freq = FALSE, main = paste("Histogram:", values$normality_results$variable), xlab = values$normality_results$variable, col = "#5e72e4")
      curve(dnorm(x, mean = mean(var_data), sd = sd(var_data)), add = TRUE, col = "#f5365c", lwd = 2)
      qqnorm(var_data, main = "Q-Q Plot Normal", col = "#5e72e4"); qqline(var_data, col = "#f5365c", lwd = 2)
      dev.off()
    }
  )
  
  output$download_norm_interpretation <- downloadHandler(
    filename = function() paste0("Interpretasi_Normalitas_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$normality_results)
      p_values <- sapply(values$normality_results$tests, function(x) x$p.value)
      normal_count <- sum(p_values > values$normality_results$alpha)
      total_tests <- length(p_values)
      
      kesimpulan <- if(normal_count >= total_tests/2) {
        "Data kemungkinan berdistribusi normal. Uji parametrik dapat digunakan."
      } else {
        "Data kemungkinan tidak berdistribusi normal. Pertimbangkan uji non-parametrik atau transformasi data."
      }
      
      content <- paste(
        "KESIMPULAN KESELURUHAN:\n",
        "=======================\n",
        "Uji yang mendukung normalitas:", normal_count, "dari", total_tests, "uji\n\n",
        "KESIMPULAN:", kesimpulan
      )
      doc <- create_word_doc("Interpretasi Uji Normalitas", content)
      print(doc, target = file)
    }
  )
  
  output$download_homo_results <- downloadHandler(
    filename = function() paste0("Hasil_Uji_Homogenitas_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$homogeneity_results)
      
      doc <- read_docx() %>%
        body_add_par("Hasil Uji Homogenitas Varians", style = "heading 1") %>%
        body_add_par(paste("Variabel:", values$homogeneity_results$variable, "~", values$homogeneity_results$group)) %>%
        body_add_par(paste("Metode:", str_to_title(values$homogeneity_results$method))) %>%
        body_add_par("")
      
      # Capture the print output of the test
      results_text <- capture.output(print(values$homogeneity_results$test))
      for(line in results_text){
        doc <- body_add_par(doc, line, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  output$download_homo_interpretation <- downloadHandler(
    filename = function() paste0("Interpretasi_Homogenitas_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$homogeneity_results)
      
      p_val <- if(values$homogeneity_results$method == "levene") values$homogeneity_results$test$`Pr(>F)`[1] else values$homogeneity_results$test$p.value
      
      kesimpulan <- ifelse(p_val > values$homogeneity_results$alpha, 
                           "Asumsi homogenitas varians terpenuhi. Ini berarti variabilitas data konsisten di semua grup yang dibandingkan.",
                           "Asumsi homogenitas varians tidak terpenuhi. Ini berarti variabilitas data berbeda secara signifikan antar grup."
      )
      
      rekomendasi <- ifelse(p_val > values$homogeneity_results$alpha,
                            "Lanjutkan analisis menggunakan uji parametrik standar seperti ANOVA atau Independent t-test.",
                            "Gunakan uji alternatif yang tidak memerlukan asumsi ini, seperti Welch's t-test atau Welch's ANOVA."
      )
      
      content <- paste(
        "INTERPRETASI UJI HOMOGENITAS\n",
        "=============================\n\n",
        "Hipotesis Nol (H0): Varians antar grup adalah sama (homogen).\n",
        "Hipotesis Alternatif (H1): Setidaknya satu varians grup berbeda.\n\n",
        "Hasil:\n",
        "p-value =", round(p_val, 5), "\n\n",
        "Kesimpulan:\n",
        kesimpulan, "\n\n",
        "Rekomendasi:\n",
        rekomendasi
      )
      doc <- create_word_doc("Interpretasi Uji Homogenitas", content)
      print(doc, target = file)
    }
  )
  
  output$download_homo_plot_jpg <- downloadHandler(
    filename = function() paste0("Plot_Homogenitas_", Sys.Date(), ".jpg"),
    content = function(file) {
      req(values$homogeneity_results)
      jpeg(file, width = 600, height = 500, quality = 95)
      plot_data <- data.frame(value = sovi_data[[values$homogeneity_results$variable]], group = values$homogeneity_results$group_data)
      plot_data <- plot_data[complete.cases(plot_data), ]
      boxplot(value ~ group, data = plot_data, main = "Boxplot per Grup", xlab = values$homogeneity_results$group, ylab = values$homogeneity_results$variable, col = viridis(length(unique(plot_data$group))))
      dev.off()
    }
  )
  
  output$download_assumptions_report <- downloadHandler(
    filename = function() paste0("Laporan_Uji_Asumsi_", Sys.Date(), ".docx"),
    content = function(file) {
      doc <- read_docx() %>%
        body_add_par("Laporan Lengkap Uji Asumsi", style = "heading 1") %>%
        body_add_par(paste("Tanggal:", Sys.Date())) %>%
        body_add_par("")
      
      # Tambahkan hasil normalitas jika ada
      if(!is.null(values$normality_results)){
        doc <- doc %>% 
          body_add_par("Hasil Uji Normalitas", style = "heading 2") %>%
          body_add_par(paste("Variabel:", values$normality_results$variable))
        
        norm_df <- data.frame(
          Uji = names(values$normality_results$tests),
          `P-Value` = sapply(values$normality_results$tests, function(x) x$p.value)
        )
        ft_norm <- flextable(norm_df) %>% autofit()
        doc <- doc %>% body_add_flextable(ft_norm) %>% body_add_par("")
      }
      
      # Tambahkan hasil homogenitas jika ada
      if(!is.null(values$homogeneity_results)){
        doc <- doc %>% 
          body_add_par("Hasil Uji Homogenitas", style = "heading 2") %>%
          body_add_par(paste("Variabel:", values$homogeneity_results$variable, "~", values$homogeneity_results$group))
        
        p_val <- if(values$homogeneity_results$method == "levene") values$homogeneity_results$test$`Pr(>F)`[1] else values$homogeneity_results$test$p.value
        homo_df <- data.frame(
          Metode = values$homogeneity_results$method,
          `P-Value` = p_val
        )
        ft_homo <- flextable(homo_df) %>% autofit()
        doc <- doc %>% body_add_flextable(ft_homo) %>% body_add_par("")
      }
      
      # Tambahkan interpretasi gabungan
      doc <- doc %>%
        body_add_par("Interpretasi Gabungan", style = "heading 2") %>%
        body_add_par(assumptions_interpretation_reactive())
      
      print(doc, target = file)
    }
  )
  
  # ===== DOWNLOAD HANDLERS UNTUK UJI RATA-RATA =====
  
  output$download_mean_results <- downloadHandler(
    filename = function() paste0("Hasil_Uji_Rata_Rata_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$mean_test_results)
      
      doc <- read_docx()
      
      # Capture the full output from the reactive expression
      results_text <- mean_results_text_reactive()
      
      # Split the text into lines and add to the Word document
      lines <- strsplit(results_text, "\n")[[1]]
      for(line in lines){
        doc <- body_add_par(doc, line, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  output$download_mean_interpretation <- downloadHandler(
    filename = function() paste0("Interpretasi_Uji_Rata_Rata_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$mean_test_results)
      
      # Capture the full interpretation from the reactive expression
      interp_text <- mean_interpretation_text_reactive()
      
      doc <- read_docx()
      lines <- strsplit(interp_text, "\n")[[1]]
      for(line in lines){
        doc <- body_add_par(doc, line, style = "Normal")
      }
      
      print(doc, target = file)
    }
  )
  
  output$download_mean_plot_jpg <- downloadHandler(
    filename = function() paste0("Plot_Uji_Rata_Rata_", Sys.Date(), ".jpg"),
    content = function(file) {
      req(mean_test_plot_object())
      p <- mean_test_plot_object()
      ggsave(file, plot = p, device = "jpeg", width = 8, height = 6, dpi = 300)
    }
  )
  
  # ===== DOWNLOAD HANDLERS UNTUK UJI PROPORSI =====
  
  output$download_prop_results <- downloadHandler(
    filename = function() paste0("Hasil_Uji_Proporsi_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$prop_test_results)
      content <- prop_results_text_reactive()
      doc <- read_docx()
      lines <- strsplit(content, "\n")[[1]]
      for(line in lines){
        doc <- body_add_par(doc, line, style = "Normal")
      }
      print(doc, target = file)
    }
  )
  
  output$download_prop_interpretation <- downloadHandler(
    filename = function() paste0("Interpretasi_Uji_Proporsi_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$prop_test_results)
      content <- prop_interpretation_text_reactive()
      doc <- read_docx()
      lines <- strsplit(content, "/n")[[1]]
      for(line in lines){
        doc <- body_add_par(doc, line, style = "Normal")
      }
      print(doc, target = file)
    }
  )
  
  output$download_prop_plot_jpg <- downloadHandler(
    filename = function() paste0("Plot_Uji_Proporsi_", Sys.Date(), ".jpg"),
    content = function(file) {
      req(prop_test_plot_object())
      plot_func <- prop_test_plot_object()
      
      jpeg(file, width = 800, height = 600, quality = 95)
      p <- plot_func()
      if(inherits(p, "ggplot")) {
        print(p)
      }
      dev.off()
    }
  )
  
  # ===== DOWNLOAD HANDLERS UNTUK UJI VARIANS =====
  
  output$download_var_results <- downloadHandler(
    filename = function() paste0("Hasil_Uji_Varians_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$var_test_results)
      content <- capture.output(output$var_test_results())
      doc <- create_word_doc("Hasil Uji Varians", content)
      print(doc, target = file)
    }
  )
  
  output$download_var_interpretation <- downloadHandler(
    filename = function() paste0("Interpretasi_Uji_Varians_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$var_test_results)
      content <- capture.output(output$var_test_interpretation())
      doc <- create_word_doc("Interpretasi Uji Varians", content)
      print(doc, target = file)
    }
  )
  
  output$download_var_plot_jpg <- downloadHandler(
    filename = function() paste0("Plot_Uji_Varians_", Sys.Date(), ".jpg"),
    content = function(file) {
      req(values$var_test_results)
      jpeg(file, width = 800, height = 600, quality = 95)
      output$var_test_plot()
      dev.off()
    }
  )
  
  # ===== DOWNLOAD HANDLERS UNTUK ANOVA =====
  
  output$download_anova_results <- downloadHandler(
    filename = function() paste0("Hasil_ANOVA_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$anova_results)
      doc <- read_docx() %>%
        body_add_par("Hasil ANOVA", style = "heading 1") %>%
        body_add_par(paste("Tanggal:", Sys.Date())) %>%
        body_add_par("") %>%
        body_add_par("Tabel ANOVA", style = "heading 2")
      
      anova_summary <- capture.output(summary(values$anova_results$model))
      for(line in anova_summary){
        doc <- body_add_par(doc, line, style = "Normal")
      }
      
      if(!is.null(values$anova_results$posthoc)){
        doc <- doc %>% 
          body_add_par("") %>%
          body_add_par("Hasil Uji Post-hoc", style = "heading 2")
        posthoc_summary <- capture.output(print(values$anova_results$posthoc))
        for(line in posthoc_summary){
          doc <- body_add_par(doc, line, style = "Normal")
        }
      }
      
      print(doc, target = file)
    }
  )
  
  output$download_anova_plots <- downloadHandler(
    filename = function() paste0("Plot_ANOVA_", Sys.Date(), ".jpg"),
    content = function(file) {
      req(values$anova_results)
      jpeg(file, width = 1000, height = 500, quality = 95)
      output$anova_plot()
      dev.off()
    }
  )
  
  output$download_anova_interpretation <- downloadHandler(
    filename = function() paste0("Interpretasi_ANOVA_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$anova_results)
      content <- capture.output(output$anova_interpretation())
      doc <- create_word_doc("Interpretasi ANOVA", content)
      print(doc, target = file)
    }
  )
  
  # ===== DOWNLOAD HANDLERS UNTUK REGRESI =====
  
  output$download_reg_results <- downloadHandler(
    filename = function() paste0("Hasil_Regresi_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$regression_model)
      doc <- read_docx() %>%
        body_add_par("Hasil Regresi Linear Berganda", style = "heading 1")
      
      results_text <- capture.output(summary(values$regression_model$model))
      for(line in results_text){
        doc <- body_add_par(doc, line, style = "Normal")
      }
      print(doc, target = file)
    }
  )
  
  output$download_reg_diagnostics <- downloadHandler(
    filename = function() paste0("Plot_Diagnostik_Regresi_", Sys.Date(), ".jpg"),
    content = function(file) {
      req(values$regression_model)
      jpeg(file, width = 800, height = 800, quality = 95)
      par(mfrow = c(2, 2))
      plot(values$regression_model$model)
      dev.off()
    }
  )
  
  output$download_reg_interpretation <- downloadHandler(
    filename = function() paste0("Interpretasi_Regresi_", Sys.Date(), ".docx"),
    content = function(file) {
      req(values$regression_model)
      content <- capture.output(output$regression_interpretation())
      doc <- create_word_doc("Interpretasi Regresi Linear Berganda", content)
      print(doc, target = file)
    }
  )
}

# Menjalankan aplikasi
shinyApp(ui = ui, server = server)
