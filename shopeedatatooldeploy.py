import os
import io
import pandas as pd
import glob2
import matplotlib.pyplot as plt
import seaborn as sns
import requests

from msal import ConfidentialClientApplication

# # Authentication to the app

# # Environment variables
# client_id = os.getenv("CLIENT_ID")
# client_secret = os.getenv("CLIENT_SECRET")
# tenant_id = os.getenv("TENANT_ID")
# site_id = os.getenv("SITE_ID")  # The SharePoint site ID
# drive_id = os.getenv("DRIVE_ID")  # The document library ID (Drive ID)
# folder_order_id = os.getenv(
#     "FOLDER_ID"
# )  # The Order All folder ID within the document library
# folder_income_id = os.getenv(
#     "FOLDER_ID"
# )  # The Income All folder ID within the document library

# # Authenticate with Microsoft Graph
# authority_url = f"https://login.microsoftonline.com/{tenant_id}"
# scopes = ["https://graph.microsoft.com/.default"]

# app = ConfidentialClientApplication(
#     client_id,
#     authority=authority_url,
#     client_credential=client_secret,
# )

# result = app.acquire_token_for_client(scopes=scopes)

# if "access_token" in result:
#     access_token = result["access_token"]
# else:
#     raise Exception("Could not acquire access token")

# headers = {
#     "Authorization": f"Bearer {access_token}",
#     "Content-Type": "application/json",
# }


# # Function to list all order files in a SharePoint folder
# def sharepoint_list_order_files():
#     url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_order_id}/children"
#     response = requests.get(url, headers=headers)
#     response.raise_for_status()
#     return response.json().get("value", [])


# def sharepoint_list_income_files():
#     url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{folder_income_id}/children"
#     response = requests.get(url, headers=headers)
#     response.raise_for_status()
#     return response.json().get("value", [])


# # Function to read a file from SharePoint into a DataFrame
# def read_file_from_sharepoint(file_id):
#     download_url = f"https://graph.microsoft.com/v1.0/sites/{site_id}/drives/{drive_id}/items/{file_id}/content"
#     response = requests.get(download_url, headers=headers)
#     response.raise_for_status()
#     file_content = response.content
#     df = pd.read_excel(io.BytesIO(file_content))
#     return df


# if __name__ == "__main__":
#     result_df = main_function()
#     print(result_df.head())


# # Ganti dengan path folder yang sesuai
# folder_path = "../FiSystem/"

# # Buat pola untuk mencari file yang sesuai
# pattern_order = folder_path + "order/*"
# pattern_income = folder_path + "income/*"

# # Buat daftar semua file Excel yang sesuai dengan pola
# file_list_order = glob2.glob(pattern_order)
# file_list_income = glob2.glob(pattern_income)

# # Buat list untuk menyimpan setiap DataFrame
# df_list_order = []
# df_list_income = []

# # Baca setiap file dan tambahkan ke list
# for file in file_list_order:
#     try:
#         df = pd.read_excel(file, dtype=str)
#         df_list_order.append(df)
#     except Exception as e:
#         print(f"Error reading {file}: {e}")


# # Gabungkan semua DataFrame dalam listfor file in file_list_income:
# for file in file_list_income:
#     try:
#         # Baca semua sheet di file Excel
#         xls = pd.ExcelFile(file)
#         # Iterasi melalui semua sheet
#         for sheet_name in xls.sheet_names:
#             # Cek apakah nama sheet mengandung "Income"
#             if "Income" in sheet_name:
#                 df = pd.read_excel(file, sheet_name=sheet_name, skiprows=5)
#                 df_list_income.append(df)
#     except Exception as e:
#         print(f"Error reading {file} or sheet {sheet_name}: {e}")

# try:
#     df_order = pd.concat(df_list_order, ignore_index=True)
# except Exception as e:
#     print(f"Error concatenating order dataframes: {e}")

# try:
#     df_income = pd.concat(df_list_income, ignore_index=True)
# except Exception as e:
#     print(f"Error concatenating income dataframes: {e}")


# try:
#     jumlah_no_pesanan_unik = df_order.drop_duplicates(
#         subset="No. Pesanan", keep="first"
#     ).shape
# except Exception as e:
#     print(f"An error occurred while dropping duplicates in jumlah_no_pesanan_unik: {e}")
# try:
#     df_order["Waktu Pesanan Dibuat"] = pd.to_datetime(df_order["Waktu Pesanan Dibuat"])
#     df_order["tahun"] = df_order["Waktu Pesanan Dibuat"].dt.year
#     df_order["bulan"] = df_order["Waktu Pesanan Dibuat"].dt.month
# except Exception as e:
#     print(
#         f"An error occurred while processing datetime columns in df_order['Waktu Pesanan Dibuat'] : {e}"
#     )
# try:
#     df_order["Total Harga Produk"] = (
#         df_order["Total Harga Produk"].astype(str).str.replace(".", "").astype(float)
#     )
# except Exception as e:
#     print(f"An error occurred while processing 'Total Harga Produk': {e}")
# try:
#     agg_order = (
#         df_order.groupby(["tahun", "bulan"])
#         .agg(qty=("No. Pesanan", "nunique"), total_harga=("Total Harga Produk", "sum"))
#         .reset_index()
#     )
# except Exception as e:
#     print(f"An error occurred while aggregating data in agg_order: {e}")
# try:
#     agg_order
# except Exception as e:
#     print("Agg_Order Failed")
# try:
#     df_order_batal = df_order[df_order["Status Pesanan"] == "Batal"]
# except Exception as e:
#     print(f"An error occurred while filtering in df_order_batal: {e}")
# try:
#     agg_batal = (
#         df_order_batal.groupby(["tahun", "bulan"])
#         .agg(qty=("No. Pesanan", "nunique"), total_harga=("Total Harga Produk", "sum"))
#         .reset_index()
#     )
# except Exception as e:
#     print(f"An error occurred while aggregating data in agg_batal: {e}")
# try:
#     df_order_sedang_dikirim = df_order[df_order["Status Pesanan"] == "Sedang Dikirim"]
# except Exception as e:
#     print(f"An error occurred while filtering in df_order_batal: {e}")
# try:
#     agg_piutang = (
#         df_order_sedang_dikirim.groupby(["tahun", "bulan"])
#         .agg(
#             qty=("No. Pesanan", "nunique"),
#             total_harga=("Total Harga Produk", "sum"),
#         )
#         .reset_index()
#     )
# except Exception as e:
#     print(f"An error occurred while aggregating data in agg_piutang: {e}")
# try:
#     agg_piutang
# except Exception as e:
#     print("Agg Piutang not found, data is empty")
# try:
#     df_income["Waktu Pesanan Dibuat"] = pd.to_datetime(
#         df_income["Waktu Pesanan Dibuat"]
#     )
#     df_income["tahun"] = df_income["Waktu Pesanan Dibuat"].dt.year
#     df_income["bulan"] = df_income["Waktu Pesanan Dibuat"].dt.month
# except Exception as e:
#     print(
#         f"An error occurred while processing datetime columns in df_income['Waktu Pesanan Dibuat'] : {e}"
#     )

# # Harus diganti agar data bisa dinamis
# df_income = df_income[df_income["tahun"] >= 2023]
# df_income = df_income[df_income["bulan"] == 7]
# try:
#     agg_income = (
#         df_income.groupby(["tahun", "bulan"])
#         .agg(
#             qty=("No. Pesanan", "nunique"),
#             total_harga_sebelum_diskon=("Harga Asli Produk", "sum"),
#             total_diskon=("Total Diskon Produk", "sum"),
#         )
#         .reset_index()
#     )
# except Exception as e:
#     print(f"An error occurred while aggregating data in agg_income: {e}")

# agg_income
# try:
#     agg_income["total_harga_setelah_diskon"] = (
#         agg_income["total_harga_sebelum_diskon"] + agg_income["total_diskon"]
#     )
# except Exception as e:
#     print(f"An error occurred while adding 'total_harga_setelah_diskon' column: {e}")

# agg_income
# try:
#     jumlah_pesanan_selesai = agg_order.merge(
#         agg_batal, on=["tahun", "bulan"], suffixes=("", "_batal")
#     )
# except Exception as e:
#     print(f"An error occurred while merging 'agg_order' and 'agg_batal': {e}")

# try:
#     jumlah_pesanan_selesai["qty"] = (
#         jumlah_pesanan_selesai["qty"] - jumlah_pesanan_selesai["qty_batal"]
#     )
#     jumlah_pesanan_selesai["total_harga"] = (
#         jumlah_pesanan_selesai["total_harga"]
#         - jumlah_pesanan_selesai["total_harga_batal"]
#     )
# except Exception as e:
#     print(f"An error occurred while subtracting in jumlah_pesanan_selesai: {e}")
# try:
#     jumlah_pesanan_selesai = jumlah_pesanan_selesai.drop(
#         columns=["qty_batal", "total_harga_batal"]
#     )
# except Exception as e:
#     print(f"An error occurred while dropping columns in jumlah_pesanan_selesai: {e}")
# try:
#     df_income["Tanggal Dana Dilepaskan"] = pd.to_datetime(
#         df_income["Tanggal Dana Dilepaskan"]
#     )
# except Exception as e:
#     print(
#         f"An error occurred while processing datetime columns in df_income['Tanggal Dana Dilepaskan'] : {e}"
#     )
# try:
#     df_income["bulan_ke"] = (
#         df_income["Tanggal Dana Dilepaskan"].dt.month - df_income["bulan"] + 1
#     )
#     # Replace values greater than 3 with '>3' perlu data yg sesuai baru bisa di coba
#     # df_income["bulan_ke"] = df_income["bulan_ke"].apply(lambda x: '>3' if x > 3 else x)

# except Exception as e:
#     print(f"An error occurred while adding 'bulan_ke' column: {e}")

# try:
#     agg_bulan = (
#         df_income.groupby(["tahun", "bulan_ke"])
#         .agg(
#             qty=("No. Pesanan", "nunique"),
#             total_harga_sebelum_diskon=("Harga Asli Produk", "sum"),
#             total_diskon=("Total Diskon Produk", "sum"),
#         )
#         .reset_index()
#     )
# except Exception as e:
#     print(f"An error occurred while aggregating data in agg_bulan: {e}")
# agg_bulan
# try:
#     agg_bulan["total_harga_setelah_diskon"] = (
#         agg_bulan["total_harga_sebelum_diskon"] + agg_bulan["total_diskon"]
#     )
# except Exception as e:
#     print(f"An error occurred while adding 'total_harga_setelah_diskon' column: {e}")

# try:
#     persentase = jumlah_pesanan_selesai.merge(
#         agg_batal, on=["tahun", "bulan"], suffixes=("", "_batal")
#     )
# except Exception as e:
#     print(
#         f"An error occurred while merging 'jumlah_pesanan_selesai' and 'agg_batal': {e}"
#     )

# try:
#     persentase["persentase_selesai"] = persentase["qty"] / (
#         persentase["qty"] + persentase["qty_batal"]
#     )
#     persentase["persentase_batal"] = persentase["qty_batal"] / (
#         persentase["qty"] + persentase["qty_batal"]
#     )
# except Exception as e:
#     print(f"An error occurred while calculating percentages in persentase: {e}")

# # persentase["persentase_selesai"], persentase["persentase_batal"] = persentase["persentase_selesai"].round().astype(int).astype(str) + '%', persentase["persentase_batal"].round().astype(int).astype(str) + '%'
# try:
#     agg_refund = (
#         df_income.groupby(["tahun", "bulan_ke"])
#         .agg(refund_total=("Jumlah Pengembalian Dana ke Pembeli", "sum"))
#         .reset_index()
#     )
# except Exception as e:
#     print(f"An error occurred while aggregating data in agg_refund: {e}")
# try:
#     agg_pendapatan_lainnya = (
#         df_income.groupby(["tahun", "bulan_ke"])
#         .agg(
#             a1=("Diskon Produk dari Shopee", "sum"),
#             a2=("Ongkir Dibayar Pembeli", "sum"),
#             a3=("Diskon Ongkir Ditanggung Jasa Kirim", "sum"),
#             a4=("Gratis Ongkir dari Shopee", "sum"),
#         )
#         .reset_index()
#     )
# except Exception as e:
#     print(f"An error occurred while aggregating data in agg_pendapatan_lainnya: {e}")
# try:
#     agg_pendapatan_lainnya["pendapatan_lainnya_total"] = (
#         agg_pendapatan_lainnya["a1"]
#         + agg_pendapatan_lainnya["a2"]
#         + agg_pendapatan_lainnya["a3"]
#         + agg_pendapatan_lainnya["a4"]
#     )
# except Exception as e:
#     print(f"An error occurred while calculating 'pendapatan_lainnya_total': {e}")
# try:
#     agg_pendapatan_lainnya = agg_pendapatan_lainnya.drop(
#         columns=["a1", "a2", "a3", "a4"]
#     )
# except Exception as e:
#     print(f"An error occurred while dropping columns in agg_pendapatan_lainnya: {e}")
# try:
#     agg_promo_toko = (
#         df_income.groupby(["tahun", "bulan_ke"])
#         .agg(
#             a1=("Diskon Voucher Ditanggung Penjual", "sum"),
#             a2=("Cashback Koin yang Ditanggung Penjual", "sum"),
#         )
#         .reset_index()
#     )
# except Exception as e:
#     print(f"An error occurred while aggregating data in agg_promo_toko: {e}")
# try:
#     agg_promo_toko["promo_toko_total"] = agg_promo_toko["a1"] + agg_promo_toko["a2"]
# except Exception as e:
#     print(f"An error occurred while calculating 'promo_toko_total': {e}")
# try:
#     agg_promo_toko = agg_promo_toko.drop(columns=["a1", "a2"])
# except Exception as e:
#     print(f"An error occurred while dropping columns in agg_promo_toko: {e}")
# try:
#     agg_potongan_lainnya = (
#         df_income.groupby(["tahun", "bulan_ke"])
#         .agg(
#             a1=("Ongkir yang Diteruskan oleh Shopee ke Jasa Kirim", "sum"),
#             a2=("Ongkos Kirim Pengembalian Barang", "sum"),
#         )
#         .reset_index()
#     )
# except Exception as e:
#     print(f"An error occurred while aggregating data in agg_potongan_lainnya: {e}")
# try:
#     agg_potongan_lainnya["potongan_lainnya_total"] = (
#         agg_potongan_lainnya["a1"] + agg_potongan_lainnya["a2"]
#     )
# except Exception as e:
#     print(f"An error occurred while calculating 'potongan_lainnya_total': {e}")
# try:
#     agg_potongan_lainnya = agg_potongan_lainnya.drop(columns=["a1", "a2"])
# except Exception as e:
#     print(f"An error occurred while dropping columns in agg_potongan_lainnya: {e}")
# try:
#     agg_biaya_marketplace = (
#         df_income.groupby(["tahun", "bulan_ke"])
#         .agg(
#             a1=("Biaya Administrasi", "sum"),
#             a2=("Biaya Layanan (termasuk PPN 11%)", "sum"),
#             # a3=("Premi", "sum"),
#             # a4=("Biaya Transaksi", "sum"),
#             a5=("Biaya Kampanye", "sum"),
#             a6=("Bea Masuk, PPN & PPh", "sum"),
#         )
#         .reset_index()
#     )
# except Exception as e:
#     print(f"An error occurred while aggregating data in agg_biaya_marketplace: {e}")
# try:
#     # agg_biaya_marketplace["biaya_market_place_total"] = agg_biaya_marketplace["a1"] + agg_biaya_marketplace["a2"] + agg_biaya_marketplace["a3"] + agg_biaya_marketplace["a4"] + agg_biaya_marketplace["a5"] + agg_biaya_marketplace["a6"]
#     agg_biaya_marketplace["biaya_market_place_total"] = (
#         agg_biaya_marketplace["a1"]
#         + agg_biaya_marketplace["a2"]
#         + agg_biaya_marketplace["a5"]
#         + agg_biaya_marketplace["a6"]
#     )
# except Exception as e:
#     print(f"An error occurred while calculating 'biaya_market_place_total': {e}")
# try:
#     agg_biaya_marketplace = agg_biaya_marketplace.drop(columns=["a1", "a2", "a5", "a6"])
# except Exception as e:
#     print(f"An error occurred while dropping columns in agg_biaya_marketplace: {e}")
# agg_biaya_marketplace
# try:
#     agg_total_released_amount = (
#         df_income.groupby(["tahun", "bulan_ke"])
#         .agg(
#             total_released_amount=("Total Penghasilan", "sum"),
#         )
#         .reset_index()
#     )
# except Exception as e:
#     print(f"An error occurred while aggregating data in agg_total_released_amount: {e}")
# try:
#     biaya_lainnya = (
#         agg_total_released_amount.merge(
#             agg_pendapatan_lainnya, on=["tahun", "bulan_ke"]
#         )
#         .merge(agg_refund, on=["tahun", "bulan_ke"])
#         .merge(agg_biaya_marketplace, on=["tahun", "bulan_ke"])
#         .merge(agg_promo_toko, on=["tahun", "bulan_ke"])
#         .merge(agg_potongan_lainnya, on=["tahun", "bulan_ke"])
#         .merge(agg_bulan, on=["tahun", "bulan_ke"])
#     )
# except Exception as e:
#     print(f"An error occurred while merging dataframes: {e}")

# try:
#     biaya_lainnya["biaya_lainnya_total"] = biaya_lainnya[
#         [
#             "pendapatan_lainnya_total",
#             "refund_total",
#             "biaya_market_place_total",
#             "promo_toko_total",
#             "total_harga_setelah_diskon",
#             "potongan_lainnya_total",
#         ]
#     ].sum(axis=1)
# except Exception as e:
#     print(f"An error occurred while calculating 'biaya_lainnya_total': {e}")
# try:
#     biaya_lainnya["biaya_lainnya_total"] = (
#         biaya_lainnya["total_released_amount"] - biaya_lainnya["biaya_lainnya_total"]
#     )
# except Exception as e:
#     print(f"An error occurred while calculating 'biaya_lainnya_total': {e}")
# try:
#     biaya_lainnya = biaya_lainnya.drop(
#         columns=["qty", "total_harga_sebelum_diskon", "total_diskon"]
#     )
# except Exception as e:
#     print(f"An error occurred while dropping columns in biaya_lainnya: {e}")
# try:
#     piutang_to_penjualan = agg_bulan.copy()
# except Exception as e:
#     print(f"An error occurred while copying DataFrame piutang_to_penjualan: {e}")

# try:
#     piutang_to_penjualan["percentage"] = (
#         piutang_to_penjualan["total_harga_setelah_diskon"]
#         / piutang_to_penjualan["total_harga_setelah_diskon"].sum()
#     )
# except Exception as e:
#     print(
#         f"An error occurred while calculating 'percentage' in piutang_to_penjualan: {e}"
#     )
# try:
#     piutang_to_penjualan = piutang_to_penjualan.rename(
#         columns={
#             "bulan_ke": "Piutang to Penjualan",
#             "total_harga_setelah_diskon": "",
#             "qty": "Qty",
#             "percentage": "Percentage",
#         }
#     )
# except Exception as e:
#     print(f"An error occurred while renaming columns in piutang_to_penjualan: {e}")
# # piutang_to_penjualan["percentage"] = piutang_to_penjualan["percentage"].round().astype(int) / 100
# piutang_to_penjualan
# try:
#     piutang_to_penjualan["Piutang to Penjualan"] = piutang_to_penjualan[
#         "Piutang to Penjualan"
#     ].apply(lambda x: f"Piutang Bulan {x}")
# except Exception as e:
#     print(
#         f"An error occurred while applying lambda function in piutang_to_penjualan['Piutang to Penjualan'] : {e}"
#     )

# try:
#     agg_bulan = agg_bulan.drop(
#         columns=["tahun", "total_harga_sebelum_diskon", "total_diskon"]
#     )
# except Exception as e:
#     print(f"An error occurred while dropping columns in agg_bulan: {e}")
# try:
#     agg_bulan = agg_bulan.rename(
#         columns={
#             "bulan_ke": "Bulan",
#             "qty": "Qty",
#             "total_harga_setelah_diskon": "Tot Harga Produk",
#         }
#     )
# except Exception as e:
#     print(f"An error occurred while renaming columns in agg_bulan: {e}")
# agg_bulan
# try:
#     agg_bulan["Bulan"] = agg_bulan["Bulan"].apply(lambda x: f"Bulan {x}")
# except Exception as e:
#     print(
#         f"An error occurred while applying lambda function in agg_bulan['Bulan'] : {e}"
#     )

# try:
#     dana_dilepas = biaya_lainnya.copy()
# except Exception as e:
#     print(f"An error occurred while copying DataFrame in dana_dilepas: {e}")

# try:
#     dana_dilepas = dana_dilepas.drop(columns="tahun")
# except Exception as e:
#     print(f"An error occurred while dropping column in dana_dilepas: {e}")
# try:
#     dana_dilepas = dana_dilepas.rename(
#         columns={
#             "bulan_ke": "Bulan",
#             "total_released_amount": "Total Released Amount",
#             "pendapatan_lainnya_total": "Pendapatan Lainnya",
#             "refund_total": "Refund",
#             "biaya_market_place_total": "Biaya Marketplace",
#             "promo_toko_total": "Promo Toko",
#             "potongan_lainnya_total": "Potongan Lainnya",
#             "total_harga_setelah_diskon": "Harga Produk",
#             "biaya_lainnya_total": "Biaya Lainnya",
#         }
#     )
# except Exception as e:
#     print(f"An error occurred while renaming columns in dana_dilepas: {e}")
# new_order_dana_dilepas = [
#     "Bulan",
#     "Pendapatan Lainnya",
#     "Promo Toko",
#     "Potongan Lainnya",
#     "Biaya Marketplace",
#     "Harga Produk",
#     "Refund",
#     "Total Released Amount",
#     "Biaya Lainnya",
# ]
# try:
#     dana_dilepas = dana_dilepas.reindex(columns=new_order_dana_dilepas)
# except Exception as e:
#     print(
#         f"An error occurred while reindexing columns in dana_dilepas: {type(e).__name__} : {e}"
#     )

# dana_dilepas
# try:
#     dana_dilepas["Bulan"] = dana_dilepas["Bulan"].apply(lambda x: f"Bulan {x}")
# except Exception as e:
#     print(
#         f"An error occurred while applying lambda function in dana_dilepas: {type(e).__name__} : {e}"
#     )

# try:
#     df_piutang = df_order.copy()
# except Exception as e:
#     print(
#         f"An error occurred while copying DataFrame df_piutang: {type(e).__name__} : {e}"
#     )

# try:
#     df_piutang = df_piutang[df_piutang["Status Pesanan"] == "Sedang Dikirim"]
# except Exception as e:
#     print(
#         f"An error occurred while filtering DataFrame in df_piutang: {type(e).__name__} : {e}"
#     )
# selected_columns = [
#     "No. Pesanan",
#     "Status Pesanan",
#     "Total Harga Produk",
#     "Jumlah",
#     "Waktu Pesanan Dibuat",
#     "Waktu Pembayaran Dilakukan",
# ]

# try:
#     df_piutang = df_piutang[selected_columns]
# except Exception as e:
#     print(
#         f"An error occurred while selecting columns in df_piutang: {type(e).__name__} : {e}"
#     )

# try:
#     df_piutang = df_piutang.reset_index(drop=True)
# except Exception as e:
#     print(
#         f"An error occurred while resetting index in df_piutang: {type(e).__name__} : {e}"
#     )

# try:
#     df_piutang["Jumlah"] = df_piutang["Jumlah"].astype(int)
# except Exception as e:
#     print(
#         f"An error occurred while converting column to integer in df_piutang: {type(e).__name__} : {e}"
#     )

# try:
#     df_piutang["Waktu Pembayaran Dilakukan"] = pd.to_datetime(
#         df_piutang["Waktu Pembayaran Dilakukan"]
#     )
# except Exception as e:
#     print(
#         f"An error occurred while converting column to datetime in df_piutang: {type(e).__name__} : {e}"
#     )

# import io


# def save_plot_to_buffer(fig):
#     buffer = io.BytesIO()
#     fig.savefig(buffer, format="png", bbox_inches="tight", pad_inches=0.1)
#     buffer.seek(0)
#     return buffer


# try:
#     qty = persentase["qty"].sum()
#     qty_batal = persentase["qty_batal"].sum()
#     total_orders_1 = qty + qty_batal
#     percentage_completed_1 = (qty / total_orders_1) * 100
#     percentage_canceled_1 = (qty_batal / total_orders_1) * 100

#     background_color = "white"
#     top_banner_color = "#D06A76"
#     section_background_color = "#FAD3D8"
#     text_color = "black"
#     gauge_color_completed = "#D06A76"
#     gauge_color_canceled = "#FF7F0E"
#     gauge_color_remaining = "#636363"

#     # Create the figure
#     fig = plt.figure(figsize=(10, 12))
#     fig.patch.set_facecolor(background_color)

#     # Top Banner
#     ax_top = plt.axes([0.05, 0.85, 0.9, 0.1])
#     ax_top.text(
#         0.05,
#         0.5,
#         "Pesanan Dibuat",
#         fontsize=18,
#         color=text_color,
#         va="center",
#         weight="bold",
#     )
#     ax_top.text(
#         0.95,
#         0.5,
#         f'{agg_order["total_harga"][0]:,.0f}\n{agg_order["qty"][0]:,.0f}',
#         fontsize=18,
#         color=text_color,
#         ha="right",
#         va="center",
#         weight="bold",
#     )
#     ax_top.set_facecolor(top_banner_color)
#     ax_top.axis("off")

#     # Canceled Orders Section
#     ax_canceled = plt.axes([0.05, 0.7, 0.4, 0.1])
#     ax_canceled.text(0.05, 0.5, "Batal", fontsize=16, color=text_color, va="center")
#     ax_canceled.text(
#         0.95,
#         0.5,
#         f'{agg_batal["total_harga"][0]:,.0f}\n{agg_batal["qty"][0]:,.0f}',
#         fontsize=16,
#         color=text_color,
#         ha="right",
#         va="center",
#     )
#     ax_canceled.set_facecolor(section_background_color)
#     ax_canceled.axis("off")

#     # Completed Orders Section
#     ax_completed = plt.axes([0.05, 0.55, 0.4, 0.1])
#     ax_completed.text(
#         0.05, 0.5, "Pesanan Selesai", fontsize=16, color=text_color, va="center"
#     )
#     ax_completed.text(
#         0.95,
#         0.5,
#         f'{jumlah_pesanan_selesai["total_harga"][0]:,.0f}\n{jumlah_pesanan_selesai["qty"][0]:,.0f}',
#         fontsize=16,
#         color=text_color,
#         ha="right",
#         va="center",
#     )
#     ax_completed.set_facecolor(section_background_color)
#     ax_completed.axis("off")

#     # Gauge Plot
#     ax_gauge = plt.axes([0.55, 0.5, 0.4, 0.4])
#     size = 0.3
#     vals = [percentage_completed_1, percentage_canceled_1]
#     colors = [gauge_color_completed, gauge_color_canceled]
#     ax_gauge.pie(
#         vals, radius=1.0, colors=colors, wedgeprops=dict(width=size, edgecolor="white")
#     )
#     ax_gauge.text(
#         0,
#         0,
#         f"{percentage_completed_1:,.0f}%",
#         fontsize=18,
#         color=text_color,
#         ha="center",
#         va="center",
#     )
#     ax_gauge.axis("equal")

#     # Bottom Text Banner
#     ax_bottom = plt.axes([0.05, 0.5, 0.9, 0.05])
#     ax_bottom.text(
#         0.77,
#         0.5,
#         "Pesanan Selesai - Batal",
#         fontsize=16,
#         color=text_color,
#         ha="center",
#         va="center",
#     )
#     ax_bottom.set_facecolor("#E37047")
#     ax_bottom.axis("off")

#     # plt.subplots_adjust(left=0.05, right=0.95, top=0.95, bottom=0.05)
#     # plt.tight_layout()
#     plt.show()

#     # Save the plot to buffer
#     buffer1 = save_plot_to_buffer(fig)
#     print("Plot saved to buffer successfully")

# except Exception as e:
#     print(f"An error occurred in Pesanan Dibuat Graph: {type(e).__name__} : {e}")

# df_piutang_grp = piutang_to_penjualan.copy()
# df_piutang_grp = df_piutang_grp.drop(columns="tahun")

# df_piutang_grp["Piutang to Penjualan"] = df_piutang_grp[
#     "Piutang to Penjualan"
# ].str.replace("Piutang ", "")
# df_piutang_grp
# if len(df_piutang_grp["Piutang to Penjualan"]) == 1:

#     df_piutang_grp.loc[len(df_piutang_grp)] = ["Bulan 2", 0, 0, 0, 0, 0]
#     df_piutang_grp.loc[len(df_piutang_grp)] = ["Bulan 3", 0, 0, 0, 0, 0]
#     df_piutang_grp.loc[len(df_piutang_grp)] = ["Bulan >3", 0, 0, 0, 0, 0]

# elif len(df_piutang_grp["Piutang to Penjualan"]) == 2:
#     df_piutang_grp.loc[len(df_piutang_grp)] = [
#         "Bulan 3",
#         0,
#         0,
#         0,
#         0,
#         0,
#     ]
#     df_piutang_grp.loc[len(df_piutang_grp)] = ["Bulan >3", 0, 0, 0, 0, 0]


# elif len(df_piutang_grp["Piutang to Penjualan"]) == 3:
#     df_piutang_grp.loc[len(df_piutang_grp)] = ["Bulan >3", 0, 0, 0, 0, 0]

# try:
#     title = "Piutang"
#     try:
#         total_prices_2 = agg_piutang["total_harga"][0]
#     except KeyError:
#         total_prices_2 = 0
#     try:
#         total_items_2 = agg_piutang["qty"][0]
#     except KeyError:
#         total_items_2 = 0
#     total_all = agg_bulan["Qty"].sum()
#     percentage_completed_2 = (total_all / (total_items_2 + total_all)) * 100

#     # Colors
#     background_color = "white"
#     top_banner_color = "#D06A76"
#     section_background_color = "#FAD3D8"
#     text_color = "black"
#     gauge_color_completed = "#D06A76"
#     gauge_color_canceled = "#FF7F0E"
#     gauge_color_remaining = "#636363"

#     # Create the figure
#     fig = plt.figure(figsize=(10, 12))
#     fig.patch.set_facecolor(background_color)

#     # Top Title Banner
#     ax_title = plt.axes([0.05, 0.85, 0.9, 0.1])
#     ax_title.text(
#         0.05, 0.5, title, fontsize=18, color="black", va="center", weight="bold"
#     )
#     ax_title.text(
#         0.95,
#         0.5,
#         f"{total_prices_2:,.0f}\n{total_items_2:,.0f}",
#         fontsize=18,
#         color="black",
#         ha="right",
#         va="center",
#         weight="bold",
#     )
#     ax_title.set_facecolor(top_banner_color)
#     ax_title.axis("off")

#     # Piutang to Penjualan Data
#     for i, row in df_piutang_grp.iterrows():
#         y_position = 0.70 - i * 0.1  # Adjust this value to change spacing
#         ax_bulan = plt.axes([0.05, y_position, 0.4, 0.1])
#         ax_bulan.text(
#             0.05,
#             0.6,
#             f'{row["Piutang to Penjualan"]}',
#             fontsize=16,
#             color=text_color,
#             fontweight="bold",
#             va="center",
#         )
#         ax_bulan.text(
#             0.05,
#             0.3,
#             f'{row["Percentage"] * 100:.1f}%',
#             fontsize=16,
#             color=text_color,
#             va="center",
#         )
#         ax_bulan.text(
#             0.95,
#             0.6,
#             f'{row[""]:,.0f}\n{row["Qty"]}',
#             fontsize=16,
#             color=text_color,
#             ha="right",
#             va="center",
#         )
#         ax_bulan.set_facecolor(section_background_color)
#         ax_bulan.axis("off")

#     # Gauge Plot
#     ax_gauge = plt.axes([0.55, 0.45, 0.4, 0.4])
#     size = 0.3
#     vals = [total_items_2, total_all - total_items_2]
#     colors = [gauge_color_canceled, gauge_color_completed]
#     ax_gauge.pie(
#         vals, radius=1.0, colors=colors, wedgeprops=dict(width=size, edgecolor="white")
#     )
#     ax_gauge.text(
#         0,
#         0,
#         f"{percentage_completed_2:,.3f}%",
#         fontsize=18,
#         color="black",
#         ha="center",
#         va="center",
#     )
#     ax_gauge.axis("equal")

#     # Bottom Text Banner
#     ax_bottom = plt.axes([0.06, 0.35, 0.9, 0.25])
#     ax_bottom.text(
#         0.77,
#         0.5,
#         "Piutang Dibayar",
#         fontsize=16,
#         color="black",
#         ha="center",
#         va="center",
#     )
#     ax_bottom.set_facecolor("#E37047")
#     ax_bottom.axis("off")

#     plt.show()
#     buffer2 = save_plot_to_buffer(fig)
# except Exception as e:
#     print(f"An error occurred in Piutang Graph: {type(e).__name__} : {e}")


# try:
#     selisih = (
#         biaya_lainnya["total_released_amount"].sum()
#         - jumlah_pesanan_selesai["total_harga"][0]
#     )
#     dana_dilepas_h_produk = 90
#     pendapatan_lainnya = dana_dilepas["Pendapatan Lainnya"].sum()
#     promo_toko = dana_dilepas["Promo Toko"].sum()
#     potongan_lainnya = dana_dilepas["Potongan Lainnya"].sum()
#     biaya_marketplace = dana_dilepas["Biaya Marketplace"].astype(int).sum()
#     refund = dana_dilepas["Refund"].sum()
#     biaya_lainnya_1 = dana_dilepas["Biaya Lainnya"].astype(int).sum()

#     # Adaptable Pie chart data
#     percentage = (
#         biaya_lainnya["total_released_amount"].sum()
#         / (
#             biaya_lainnya["total_released_amount"].sum()
#             + agg_bulan["Tot Harga Produk"].sum()
#             - biaya_lainnya["total_released_amount"].sum()
#         )
#         * 100
#     )
#     remaining_percentage = 100 - dana_dilepas_h_produk
#     pie_sizes = [
#         biaya_lainnya["total_released_amount"].sum(),
#         agg_bulan["Tot Harga Produk"].sum()
#         - biaya_lainnya["total_released_amount"].sum(),
#     ]
#     pie_colors = ["#D06A76", "#FF7F0E"]  # User-friendly colors: green tones

#     # Create figure and axes with increased size
#     fig, ax = plt.subplots(figsize=(8, 10))  # Increased size for better readability

#     # Selisih text
#     plt.text(
#         0.5,
#         0.95,
#         "Selisih",
#         fontsize=16,
#         weight="bold",
#         ha="center",
#         transform=plt.gcf().transFigure,
#     )  # Adjust font size
#     plt.text(
#         0.5,
#         0.90,
#         f"{selisih:,.0f}",
#         fontsize=16,
#         weight="bold",
#         ha="center",
#         transform=plt.gcf().transFigure,
#     )  # Adjust font size

#     # Pie chart
#     ax.pie(
#         pie_sizes,
#         labels=["", ""],
#         colors=pie_colors,
#         startangle=90,
#         counterclock=False,
#         wedgeprops={"width": 1.6},
#     )
#     centre_circle = plt.Circle(
#         (0, 0), 0.60, fc="white"
#     )  # Adjust size of the center circle here
#     fig.gca().add_artist(centre_circle)
#     ax.text(
#         0,
#         0,
#         f"{percentage:,.1f}%",
#         ha="center",
#         va="center",
#         fontsize=20,
#         weight="bold",
#     )  # Adjust font size
#     ax.set_title(
#         "Dana Dilepas - H Produk", fontsize=14, weight="bold", y=0.01
#     )  # Adjust title font size

#     # Table data
#     labels = [
#         "Pendapatan Lainnya",
#         "Promo Toko",
#         "Potongan Lainnya",
#         "Biaya Marketplace",
#         "Refund",
#         "Biaya Lainnya",
#     ]
#     values = [
#         pendapatan_lainnya,
#         promo_toko,
#         potongan_lainnya,
#         biaya_marketplace,
#         refund,
#         biaya_lainnya_1,
#     ]

#     # Create table data
#     table_data = [[label, f"{value:,}"] for label, value in zip(labels, values)]

#     # Add table
#     table = plt.table(
#         cellText=table_data,
#         colLabels=["Category", "Value"],
#         loc="bottom",
#         cellLoc="center",
#         colColours=["#f8f8f8", "#f8f8f8"],
#         bbox=[0, -0.5, 1, 0.5],
#     )
#     table.auto_set_font_size(False)
#     table.set_fontsize(12)  # Adjust font size

#     # Customize table appearance
#     row_colors = ["#f0f0f0", "#ffffff"]
#     for i, key in enumerate(table.get_celld().keys()):
#         cell = table.get_celld()[key]
#         cell.set_edgecolor("white")  # Remove cell borders
#         cell.set_facecolor(row_colors[key[0] % len(row_colors)])  # Alternate row colors
#         cell.set_text_props(
#             ha="left" if key[1] == 0 else "right", weight="bold"
#         )  # Align text and adjust font size

#     # Set column header style
#     table[0, 0].set_text_props(weight="bold", color="black")  # Adjust header font size
#     table[0, 1].set_text_props(weight="bold", color="black")  # Adjust header font size
#     table[0, 0].set_facecolor("#dcdcdc")
#     table[0, 1].set_facecolor("#dcdcdc")

#     # Remove axes
#     ax.axis("equal")
#     plt.axis("off")

#     # Show the plot
#     plt.show()
#     buffer3 = save_plot_to_buffer(fig)
# except Exception as e:
#     print(f"An error occurred in Selisih Graph: {type(e).__name__} : {e}")
# import seaborn as sns

# biaya_lainnya
# df_dana_dilepas = biaya_lainnya.copy()
# dana_dilepas
# import matplotlib.ticker as ticker

# try:
#     categories = [
#         "Pendapatan Lainnya",
#         "Promo Toko",
#         "Potongan Lainnya",
#         "Biaya Marketplace",
#         "Harga Produk",
#         "Refund",
#         "Total Released Amount",
#         "Biaya Lainnya",
#     ]

#     df_dana_dilepas_1 = dana_dilepas.copy()
#     if len(df_dana_dilepas_1["Bulan"]) == 1:
#         df_dana_dilepas_1.loc[len(df_dana_dilepas_1)] = [
#             "Bulan 2",
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#         ]
#         df_dana_dilepas_1.loc[len(df_dana_dilepas_1)] = [
#             "Bulan 3",
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#         ]
#         df_dana_dilepas_1.loc[len(df_dana_dilepas_1)] = [
#             "Bulan >3",
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#         ]
#         df_dana_dilepas_1.loc[len(df_dana_dilepas_1)] = ["", 0, 0, 0, 0, 0, 0, 0, 0]
#     elif len(df_dana_dilepas_1["Bulan"]) == 2:
#         df_dana_dilepas_1.loc[len(df_dana_dilepas_1)] = [
#             "Bulan 3",
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#         ]
#         df_dana_dilepas_1.loc[len(df_dana_dilepas_1)] = [
#             "Bulan >3",
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#         ]
#         df_dana_dilepas_1.loc[len(df_dana_dilepas_1)] = ["", 0, 0, 0, 0, 0, 0, 0, 0]
#     elif len(df_dana_dilepas_1["Bulan"]) == 3:
#         df_dana_dilepas_1.loc[len(df_dana_dilepas_1)] = [
#             "Bulan >3",
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#             0,
#         ]
#         df_dana_dilepas_1.loc[len(df_dana_dilepas_1)] = ["", 0, 0, 0, 0, 0, 0, 0, 0]

#     df_dana_dilepas_1.set_index("Bulan", inplace=True)
#     df_dana_dilepas_1 = df_dana_dilepas_1.transpose()

#     # Colors and markers
#     colors = [
#         "#1f77b4",
#         "#ff7f0e",
#         "#2ca02c",
#         "#d62728",
#         "#9467bd",
#         "#8c564b",
#         "#e377c2",
#         "#7f7f7f",
#     ]
#     markers = ["o", "s", "^", "D", "x", "p", "*", "h"]

#     # Plot
#     fig, ax = plt.subplots(
#         figsize=(22, 18)
#     )  # Reduced the width of the figure by 2 points

#     # Adjust the x-axis to ensure equal spacing
#     columns = list(df_dana_dilepas_1.columns)
#     positions = range(len(columns))
#     for i, category in enumerate(categories):
#         ax.plot(
#             positions,
#             df_dana_dilepas_1.loc[category],
#             label=category,
#             marker=markers[i],
#             color=colors[i],
#             linestyle="-",
#         )

#     # Set x-ticks and labels to ensure proper spacing
#     ax.set_xticks(positions)
#     ax.set_xticklabels(columns)

#     # Set y-axis limits
#     ax.set_ylim(-600000000, 1200000000)

#     # Add labels and title
#     ax.set_title("Dana Dilepas", fontsize=30, pad=20)
#     ax.set_ylabel("Amount (IDR)", fontsize=28)
#     ax.set_xlabel("", fontsize=28)
#     ax.legend(loc="best", fontsize=20)

#     # Set grid
#     ax.grid(True, which="both", linestyle="--", linewidth=0.5)
#     ax.axhline(y=0, color="black", linewidth=1.3)

#     # Format y-axis to display numbers
#     ax.yaxis.set_major_formatter(
#         ticker.FuncFormatter(lambda x, pos: "{:,.0f}".format(x))
#     )

#     # Increase font size of tick labels
#     ax.tick_params(axis="both", which="major", labelsize=18)

#     # Top banner
#     total_amount = dana_dilepas[
#         "Total Released Amount"
#     ].sum()  # Replace with actual dynamic data
#     total_items = agg_bulan["Qty"].sum()  # Replace with actual dynamic data
#     plt.figtext(
#         0.01,
#         0.95,
#         "Total Dana Dilepas",
#         fontsize=30,
#         color="black",
#         backgroundcolor="White",
#     )
#     plt.figtext(
#         0.95,
#         0.95,
#         f"{total_amount:,.0f}\n{total_items:,.0f}",
#         fontsize=30,
#         color="black",
#         backgroundcolor="white",
#         ha="right",
#     )

#     # Table
#     df_column_dropped = df_dana_dilepas_1.drop(
#         "", axis=1
#     )  # Ensure to handle the case where '' column might not exist
#     table_data = df_column_dropped.reset_index()
#     columns = ["Category"] + df_column_dropped.columns.tolist()
#     cell_text = table_data.values.tolist()

#     # Adjust the layout to ensure proper spacing
#     plt.subplots_adjust(left=0.1, right=0.94, top=0.88, bottom=0.4)

#     # Create the table
#     chart1 = plt.table(
#         cellText=cell_text,
#         colLabels=columns,
#         cellLoc="center",
#         loc="bottom",
#         bbox=[-0.12, -0.5, 1.124, 0.4],
#     )

#     # Set font size for table cells and adjust the height
#     chart1.auto_set_font_size(False)
#     for key, cell in chart1.get_celld().items():
#         cell.set_fontsize(16)  # Set desired font size here
#         cell.set_height(0.05)  # Adjust the height for every cell

#         if key[1] == 0:  # Adjust the width for the "Category" column
#             cell.set_width(0.41)
#         else:  # Adjust the width for other columns
#             cell.set_width(0.7)
#         cell.set_text_props(ha="center", va="center")  # Center text in cells
#         if key[0] == 0:
#             cell.set_text_props(weight="bold")  # Make the header bold

#     fig.patch.set_facecolor("white")
#     plt.show()
#     buffer4 = save_plot_to_buffer(fig)
# except Exception as e:
#     df_dana_dilepas_1
#     print(f"An error occurred in Dana Dilepas Graph: {type(e).__name__} : {e}")
# from openpyxl import Workbook
# from datetime import datetime
# from openpyxl.utils.dataframe import dataframe_to_rows
# from openpyxl.drawing.image import Image
# from openpyxl.styles import NamedStyle
# from openpyxl.utils import get_column_letter
# import calendar

# try:
#     wb = Workbook()
#     ws = wb.active
#     ws.title = "Help Dashboard"

#     # Create the percentage style
#     percentage_style = NamedStyle(name="percentage_style")
#     percentage_style.number_format = "0.0%"

#     # Memasukkan judul kolom
#     ws["A2"] = "Item"
#     ws["B2"] = "Qty"
#     ws["C2"] = "Total"

#     ws["A3"] = "Jumlan Pesanan Dibuat"
#     ws["B3"] = agg_order["qty"][0]
#     ws["C3"] = agg_order["total_harga"][0]

#     ws["A5"] = "Status Batal"
#     ws["B5"] = agg_batal["qty"][0]
#     ws["C5"] = agg_batal["total_harga"][0]

#     ws["A7"] = "Jumlan Pesanan Selesai"
#     ws["B7"] = jumlah_pesanan_selesai["qty"][0]
#     ws["C7"] = jumlah_pesanan_selesai["total_harga"][0]

#     # Menentukan baris awal untuk data agg_bulan
#     start_row = (
#         ws.max_row + 2
#     )  # Tambahkan beberapa baris kosong di antara data sebelumnya

#     # Menambahkan data dari agg_bulan mulai dari start_row
#     rows = dataframe_to_rows(agg_bulan, index=False)
#     for r_idx, row in enumerate(rows, start_row):
#         for c_idx, value in enumerate(row, 1):
#             ws.cell(row=r_idx, column=c_idx, value=value)

#     start_row = start_row + len(agg_bulan) + 2
#     ws[f"A{start_row}"] = "Total Dibayar (Harga Produk)"
#     ws[f"B{start_row}"] = agg_bulan["Qty"].sum()
#     ws[f"C{start_row}"] = agg_bulan["Tot Harga Produk"].sum()

#     start_row = ws.max_row + 2
#     ws[f"A{start_row}"] = "Piutang"
#     try:
#         ws[f"B{start_row}"] = agg_piutang["qty"][0]
#     except KeyError:
#         ws[f"B{start_row}"] = 0
#     try:
#         ws[f"C{start_row}"] = agg_piutang["total_harga"][0]
#     except KeyError:
#         ws[f"C{start_row}"] = 0

#     start_row = ws.max_row + 2
#     ws[f"A{start_row}"] = "Total Potential Sales"
#     ws[f"C{start_row}"] = agg_order["total_harga"][0]
#     ws[f"A{start_row + 1}"] = "Sales"
#     ws[f"C{start_row + 1}"] = agg_bulan["Tot Harga Produk"].sum()
#     ws[f"A{start_row + 2}"] = "Lost"
#     ws[f"C{start_row + 2}"] = agg_batal["total_harga"][0]
#     ws[f"A{start_row + 3}"] = "Piutang"
#     try:
#         ws[f"B{start_row + 3}"] = agg_piutang["qty"][0]
#     except KeyError:
#         ws[f"B{start_row + 3}"] = 0
#     try:
#         ws[f"C{start_row + 3}"] = agg_piutang["total_harga"][0]
#     except KeyError:
#         ws[f"C{start_row + 3}"] = 0
#     # ws[f"B{start_row + 3}"] = agg_piutang["qty"][0]
#     # ws[f"C{start_row + 3}"] = agg_piutang["total_harga"][0]

#     start_row = ws.max_row + 3
#     ws[f"A{start_row}"] = "Pesanan Selesai to Pesanan Dibuat"
#     ws[f"C{start_row}"] = "Percentage"
#     ws[f"A{start_row + 1}"] = "Jumlah Pesanan Selesai"
#     ws[f"B{start_row + 1}"] = jumlah_pesanan_selesai["qty"][0]
#     ws[f"C{start_row + 1}"] = persentase["persentase_selesai"][0]
#     ws[f"C{start_row + 1}"].number_format = "0.0%"
#     ws[f"A{start_row + 2}"] = "Status Batal"
#     ws[f"B{start_row + 2}"] = agg_batal["qty"][0]
#     ws[f"C{start_row + 2}"] = persentase["persentase_batal"][0]
#     ws[f"C{start_row + 2}"].number_format = "0.0%"

#     start_row = ws.max_row + 2
#     rows = dataframe_to_rows(
#         piutang_to_penjualan.drop(
#             columns=["tahun", "total_diskon", "total_harga_sebelum_diskon", "Qty"]
#         ),
#         index=False,
#     )
#     for r_idx, row in enumerate(rows, start_row):
#         for c_idx, value in enumerate(row, 1):
#             ws.cell(row=r_idx, column=c_idx, value=value)

#     percentage_col_name = "Percentage"
#     percentage_col_idx = (
#         list(
#             piutang_to_penjualan.drop(
#                 columns=["tahun", "total_diskon", "total_harga_sebelum_diskon", "Qty"]
#             ).columns
#         ).index(percentage_col_name)
#         + 1
#     )

#     # Apply percentage format to the percentage column
#     for row in ws.iter_rows(
#         min_row=start_row + 1,
#         max_row=ws.max_row,
#         min_col=percentage_col_idx,
#         max_col=percentage_col_idx,
#     ):
#         for cell in row:
#             cell.style = percentage_style

#     start_row = ws.max_row + 1
#     ws[f"A{start_row}"] = "Penjualan"
#     ws[f"B{start_row}"] = agg_bulan["Tot Harga Produk"].sum()
#     ws[f"C{start_row}"] = 1 - piutang_to_penjualan["Percentage"].sum()
#     ws[f"C{start_row}"].number_format = "0.0%"

#     start_row = ws.max_row + 2
#     ws[f"A{start_row}"] = "Dana Dilepas to Harga Produk"
#     ws[f"C{start_row}"] = "Percentage"
#     ws[f"A{start_row + 1}"] = "Dana Dilepas"
#     ws[f"B{start_row + 1}"] = biaya_lainnya["total_released_amount"].sum()
#     ws[f"C{start_row + 1}"] = (
#         biaya_lainnya["total_released_amount"].sum()
#         / jumlah_pesanan_selesai["total_harga"][0]
#     )
#     ws[f"C{start_row + 1}"].number_format = "0.0%"
#     ws[f"E{start_row + 1}"] = (
#         biaya_lainnya["total_released_amount"].sum()
#         - jumlah_pesanan_selesai["total_harga"][0]
#     )
#     ws[f"A{start_row + 2}"] = "Harga Produk"
#     ws[f"B{start_row + 2}"] = jumlah_pesanan_selesai["total_harga"][0]
#     ws[f"C{start_row + 2}"] = 1 - (
#         biaya_lainnya["total_released_amount"].sum()
#         / jumlah_pesanan_selesai["total_harga"][0]
#     )
#     ws[f"C{start_row + 2}"].number_format = "0.0%"

#     start_row = ws.max_row + 3

#     # Menambahkan data dari agg_bulan mulai dari start_row
#     rows = dataframe_to_rows(dana_dilepas, index=False)
#     for r_idx, row in enumerate(rows, start_row):
#         for c_idx, value in enumerate(row, 1):
#             ws.cell(row=r_idx, column=c_idx, value=value)

#     for col in ws.columns:
#         max_length = 0
#         column = col[
#             0
#         ].column_letter  # Dapatkan nama kolom (misalnya 'A', 'B', 'C', ...)
#         for cell in col:
#             try:
#                 if len(str(cell.value)) > max_length:
#                     max_length = len(cell.value)
#             except:
#                 pass
#         adjusted_width = (max_length + 2) * 1.2  # Penyesuaian lebar dengan margin
#         ws.column_dimensions[column].width = adjusted_width

#     # Generate nama file dengan waktu hingga milidetik
#     order_time = ("-").join(
#         str(df_order["Waktu Pesanan Dibuat"][0]).split()[0].split("-")[:2]
#     )
#     current_time = datetime.now().strftime("%Y-%m-%d_%H-%M-%S-%f")[:-3]
#     file_name = f"Shopee {order_time} Update {current_time}.xlsx"


# except Exception as e:
#     print(f"An error occurred in sheet Help Dashboard: {type(e).__name__} : {e}")

# # Simpan workbook
# # wb.save(file_name)


# try:
#     ws_graphs = wb.create_sheet(title="Graphs")
#     buffers = [buffer1, buffer2, buffer3, buffer4]
#     positions = ["A1", "A30", "AX1", "P1"]  # Posisi gambar dalam sheet

#     for buffer, position in zip(buffers, positions):
#         img = Image(buffer)
#         ws_graphs.add_image(img, position)

#     # Set zoom level to 45%
#     ws_graphs.sheet_view.zoomScale = 45
#     ws_graphs.sheet_view.zoomScaleNormal = 45

# except Exception as e:
#     print(f"An error occurred in sheet Graphs: {type(e).__name__} : {e}")
# try:
#     ws_piutang = wb.create_sheet(title="Piutang")

#     rows = dataframe_to_rows(df_piutang, index=False)
#     for r_idx, row in enumerate(rows, 1):
#         for c_idx, value in enumerate(row, 1):
#             ws_piutang.cell(row=r_idx, column=c_idx, value=value)

#     for col in ws_piutang.columns:
#         max_length = 0
#         column = col[
#             0
#         ].column_letter  # Dapatkan nama kolom (misalnya 'A', 'B', 'C', ...)
#         for cell in col:
#             try:
#                 if len(str(cell.value)) > max_length:
#                     max_length = len(cell.value)
#             except:
#                 pass
#         adjusted_width = (max_length + 2) * 1.2  # Penyesuaian lebar dengan margin
#         ws_piutang.column_dimensions[column].width = adjusted_width
# # wb.save(file_name)
# except Exception as e:
#     print(f"An error occurred in sheet Piutang: {type(e).__name__} : {e}")
# try:
#     # Save the workbook
#     file_path = f"../FiSystem/dashboard/{file_name}"
#     wb.save(file_path)

# except Exception as e:
#     print(f"An error occurred while saving the file: {type(e).__name__} : {e}")

# try:
#     for buffer in buffers:
#         buffer.close()
# except Exception as e:
#     print(f"An error occurred while closing buffer: {type(e).__name__} : {e}")
