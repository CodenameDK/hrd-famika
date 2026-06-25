import * as XLSX from "xlsx";

export const exportAttendanceToExcel = ({ records, startDate, endDate ,mode}) => {
  if (!records || records.length === 0) {
    throw new Error("Tidak ada data absensi untuk diekspor");
  }

const excelData = records.map((row, index) => ({
  No: index + 1,

  // untuk unchecked: bisa range tanggal atau "-"
  Tanggal: row.isUnchecked
    ? `${startDate} s/d ${endDate}`
    : row.attendance_date,

  "Nama Karyawan": row.employee?.name || "-",
  NIK: row.employee?.nik || "-",
  "Nama PM": row.direct_pm?.name || "-",
  "Lokasi Kerja": row.employee?.work_location || "-",
  "Keterangan/Project": row.project || "-",

  "Check In": row.isUnchecked
    ? "-"
    : formatTime(row.check_in_time),

  "Check Out": row.isUnchecked
    ? "-"
    : formatTime(row.check_out_time),

  "Total Jam Kerja": row.isUnchecked
    ? "-"
    : getWorkDuration(row.check_in_time, row.check_out_time),

  Status: row.isUnchecked
    ? "Belum Check-in"
    : row.attendance_statuses?.name || "-",

  Catatan: row.notes || "",

  "Lat Check-in": row.loc_checkin?.lat ?? "-",
  "Lng Check-in": row.loc_checkin?.lng ?? "-",
  "Alamat Check-in": row.loc_checkin?.address || "-",
}));



  const worksheet = XLSX.utils.json_to_sheet(excelData);

 worksheet["!cols"] = [
  { wch: 5 },   // No
  { wch: 12 },  // Tanggal
  { wch: 25 },  // Nama Karyawan
  { wch: 15 },  // NIK
  { wch: 25 },  // Nama PM
  { wch: 18 },  // Lokasi Kerja
  { wch: 25 },  // Keterangan/Project
  { wch: 12 },  // Check In
  { wch: 12 },  // Check Out
  { wch: 18 },  // Total Jam Kerja
  { wch: 15 },  // Status
  { wch: 30 },  // Catatan
  { wch: 15 },  // Lat Check-in
  { wch: 15 },  // Lng Check-in
  { wch: 45 },  // Alamat Check-in
];


  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Laporan Absensi");
const statusLabel =
  mode === "unchecked" ? "belum-checkin" : "";
XLSX.writeFile(
  workbook,
  `laporan-absensi-${statusLabel}-${startDate}-sd-${endDate}.xlsx`
);

};

const formatTime = (iso) => {
  if (!iso) return "-";
  return new Date(iso).toLocaleTimeString("id-ID", {
    hour: "2-digit",
    minute: "2-digit",
  });
};
const getWorkDuration = (checkIn, checkOut) => {
  if (!checkIn || !checkOut) return "-";

  const inDate = new Date(checkIn);
  const outDate = new Date(checkOut);

  inDate.setSeconds(0, 0);
  outDate.setSeconds(0, 0);

  const diffMs = outDate - inDate;
  if (diffMs <= 0) return "0 jam 0 menit";

  const totalMinutes = Math.floor(diffMs / 60000);
  const hours = Math.floor(totalMinutes / 60);
  const minutes = totalMinutes % 60;

  return `${hours} jam ${minutes} menit`;
};

