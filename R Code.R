# Install dan load package yang dibutuhkan
# Menggunakan operator logical untuk memeriksa apakah package sudah diinstal
if (!require(readxl)) {
  install.packages("readxl")
  library(readxl)
}
if (!require(gridExtra)) {
  install.packages("gridExtra")
  library(gridExtra)
}
if (!require(grid)) {
  install.packages("grid")
  library(grid)
}
if (!require(png)) {
  install.packages("png")
  library(png)
}

# Membaca data dari file Excel menggunakan package readxl 
library(readxl)
dummy_data_copy_2 <- read_excel("Documents/COLLEGE/SEMESTER 6/R Programming/Mid Term/dummy_data copy 2.xlsx")
View(dummy_data_copy_2)

# Deklarasi fungsi untuk menampilkan data
tampilkan_data <- function(data) {
  print(data)
}

# Deklarasi fungsi untuk membuat sales invoice
buat_invoice <- function(data, nomor_invoice) {
  # Memfilter data berdasarkan nomor invoice menggunakan operator subsetting
  invoice_data <- data[data$Invoice_number == nomor_invoice, ]
  
  # Percabangan untuk memeriksa apakah nomor invoice ditemukan 
  if (nrow(invoice_data) == 0) {
    cat("Nomor invoice", nomor_invoice, "tidak ditemukan.\n")
    return()
  }
  
  # Perhitungan total amount menggunakan operator arithmetic
  invoice_data$total_amount <- invoice_data$quantity * invoice_data$unit_price * (1 - invoice_data$disc_percent/100)
  subtotal <- sum(invoice_data$total_amount)
  
  # Perhitungan pajak menggunakan operator arithmetic berdasarkan tax_percent
  tax_rate <- invoice_data$tax_percent[1] / 100
  tax <- subtotal * tax_rate
  total <- subtotal + tax
  
  # Membuat tabel invoice menggunakan data frame
  invoice_table <- data.frame(
    "Item Name" = invoice_data$item_name,
    "Item Description" = invoice_data$item_description,
    "Quantity" = invoice_data$quantity,
    "Unit Price" = formatC(invoice_data$unit_price, format = "f", digits = 2),
    "Disc Percent" = invoice_data$disc_percent,
    "Tax Percent" = invoice_data$tax_percent,
    "Total Amount" = formatC(invoice_data$total_amount, format = "f", digits = 2)
  )
  
  # Mengatur ukuran gambar menggunakan package png
  png(paste0("Documents/COLLEGE/SEMESTER 6/R Programming/Mid Term/invoice_", nomor_invoice, ".png"), width = 2283, height = 2954, res = 300)
  
  # Membuat layout invoice menggunakan package grid dan gridExtra
  pushViewport(viewport(layout = grid.layout(12, 10, heights = unit(c(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1), "null"))))
  
  grid.rect(gp = gpar(fill = "white"))
  grid.text("SALES INVOICE", x = 0.5, y = 0.96, gp = gpar(fontsize = 16, fontface = "bold"))
  grid.text(invoice_data$customer_name[1], x = 0.1, y = 0.92, just = "left", gp = gpar(fontsize = 12, fontface = "bold"))
  grid.text(invoice_data$invoice_date[1], x = 0.9, y = 0.92, just = "right", gp = gpar(fontsize = 12))
  grid.text(nomor_invoice, x = 0.1, y = 0.88, just = "left", gp = gpar(fontsize = 12))
  grid.text(invoice_data$customer_number[1], x = 0.9, y = 0.88, just = "right", gp = gpar(fontsize = 12))
  grid.text("Bill to", x = 0.1, y = 0.84, just = "left", gp = gpar(fontsize = 10))
  grid.text("Ship to", x = 0.55, y = 0.84, just = "left", gp = gpar(fontsize = 10))
  grid.text(invoice_data$bill_to[1], x = 0.1, y = 0.8, just = "left", gp = gpar(fontsize = 10))
  grid.text(invoice_data$ship_to[1], x = 0.55, y = 0.8, just = "left", gp = gpar(fontsize = 10))
  grid.text("Terms : ", x = 0.1, y = 0.76, just = "left", gp = gpar(fontsize = 10))
  grid.text(invoice_data$terms[1], x = 0.19, y = 0.76, just = "left", gp = gpar(fontsize = 10))
  
  pushViewport(viewport(layout.pos.row = 5, layout.pos.col = 1:10))
  grid.table(invoice_table, rows = NULL, theme = ttheme_default(base_size = 10))
  popViewport()
  
  grid.text("SUBTOTAL", x = 0.8, y = 0.3, just = "right", gp = gpar(fontsize = 10))
  grid.text(formatC(subtotal, format = "f", digits = 2), x = 0.95, y = 0.3, just = "right", gp = gpar(fontsize = 10))
  grid.text("TAX RATE", x = 0.8, y = 0.25, just = "right", gp = gpar(fontsize = 10))
  grid.text(paste(formatC(tax_rate * 100, format = "f", digits = 3), "%"), x = 0.95, y = 0.25, just = "right", gp = gpar(fontsize = 10))
  grid.text("TAX", x = 0.8, y = 0.2, just = "right", gp = gpar(fontsize = 10))
  grid.text(formatC(tax, format = "f", digits = 2), x = 0.95, y = 0.2, just = "right", gp = gpar(fontsize = 10))
  grid.text("TOTAL", x = 0.8, y = 0.15, just = "right", gp = gpar(fontsize = 10))
  grid.text(paste("$", formatC(total, format = "f", digits = 2)), x = 0.95, y = 0.15, just = "right", gp = gpar(fontsize = 12, fontface = "bold"))
  grid.text("Other Comments or Special Instructions", x = 0.1, y = 0.3, just = "left", gp = gpar(fontsize = 10))
  grid.text("1. Total payment due in 30 days", x = 0.1, y = 0.25, just = "left", gp = gpar(fontsize = 10))
  grid.text("2. Please include the invoice number on your check", x = 0.1, y = 0.2, just = "left", gp = gpar(fontsize = 10))
  
  # Menutup perangkat grafis
  dev.off()
  
  # Menampilkan pesan sukses
  cat("Invoice", nomor_invoice, "berhasil dibuat dan disimpan.\n")
}

# Menampilkan data menggunakan fungsi tampilkan_data
tampilkan_data(dummy_data_copy_2)

# Loop untuk membuat beberapa invoice
nomor_invoice_list <- c("INV001", "INV002", "INV003")

for (nomor_invoice in nomor_invoice_list) {
  # Membuat sales invoice berdasarkan nomor invoice
  buat_invoice(dummy_data_copy_2, nomor_invoice)
}