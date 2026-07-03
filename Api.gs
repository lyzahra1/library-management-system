function getBooks() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Daftar Buku");
  const data = sheet.getDataRange().getValues();
  const books = [];
  for (let i = 1; i < data.length; i++) {
    books.push({
      id: data[i][0],
      judul: data[i][1],
      penulis: data[i][2],
      kategori: data[i][3],
      stokTotal: data[i][4],
      stokTersedia: data[i][5]
    });
  }
  return books;
}

function getSummary() {
  const sheet = SpreadsheetApp.getActive().getSheetByName("Daftar Buku");
  const data = sheet.getDataRange().getValues();

  let total = 0;
  let tersedia = 0;

  for (let i = 1; i < data.length; i++) {
    total    += Number(data[i][5]) || 0; // F
    tersedia += Number(data[i][6]) || 0; // G
  }

  return {
    total,
    tersedia,
    dipinjam: total - tersedia
  };
}

function prosesPeminjaman({ ss, sheetBuku, sheetPinjam, nama, divisi, noHp, email, idBuku }) {

  // Cari baris buku berdasarkan ID
  const dataBuku = sheetBuku.getDataRange().getValues();
  let bukuRow = -1;
  let judulBuku = "";

  for (let i = 1; i < dataBuku.length; i++) {
    if (String(dataBuku[i][0]) === String(idBuku)) {
      bukuRow = i + 1;
      judulBuku = dataBuku[i][1]; // ✅ ambil judul buku
      break;
    }
  }

  if (bukuRow === -1) throw new Error("Buku tidak ditemukan: " + idBuku);

  const stokTersedia = Number(dataBuku[bukuRow - 1][5]);
  if (stokTersedia <= 0) throw new Error("Stok buku habis");

  // Kurangi stok tersedia
  sheetBuku.getRange(bukuRow, 6).setValue(stokTersedia - 1);

  // Generate ID Peminjaman
  const lastRow = sheetPinjam.getLastRow();
  const idPinjam = "PJ" + String(lastRow).padStart(4, "0");
  const tanggal = new Date();

  // ✅ Urutan sesuai kolom sheet:
  // ID Pinjam | Tanggal Pinjam | Nama | Divisi | No HP | ID Buku | Judul Buku | Tanggal Kembali | Status | Email
  sheetPinjam.appendRow([
    idPinjam,    // A - ID Pinjam
    tanggal,     // B - Tanggal Pinjam
    nama,        // C - Nama
    divisi,      // D - Divisi
    noHp,        // E - No HP
    idBuku,      // F - ID Buku
    judulBuku,   // G - Judul Buku
    "",          // H - Tanggal Kembali (kosong dulu)
    "Dipinjam",  // I - Status
    email        // J - Email ✅
  ]);

  kirimEmailPeminjaman({
  email,
  nama,
  idPinjam,
  judulBuku,
  tanggal
});
}

function prosesPengembalian({ ss, sheetBuku, sheetPinjam, sheetPengembalian, email, idPinjamInput }) {

  const dataPinjam = sheetPinjam.getDataRange().getValues();
  let pinjamRow = -1;
  let idBuku = "";

  for (let i = 1; i < dataPinjam.length; i++) {
    if (
      String(dataPinjam[i][0]) === String(idPinjamInput) &&
      String(dataPinjam[i][9]) === String(email)
    ) {
      pinjamRow = i + 1;
      idBuku = dataPinjam[i][5]; // kolom F = ID Buku
      break;
    }
  }

  if (pinjamRow === -1) throw new Error("Data peminjaman tidak ditemukan");

  const status = dataPinjam[pinjamRow - 1][8]; // kolom I = Status
  if (status === "Dikembalikan") throw new Error("Buku sudah pernah dikembalikan");

  // Ambil data dari sheet Peminjaman untuk keperluan log
  const namaPeminjam = dataPinjam[pinjamRow - 1][2]; // kolom C = Nama
  const judulBuku    = dataPinjam[pinjamRow - 1][6]; // kolom G = Judul Buku

  const tanggalKembali = new Date();

  // Update sheet Peminjaman
  sheetPinjam.getRange(pinjamRow, 8).setValue(tanggalKembali); // kolom H = Tanggal Kembali
  sheetPinjam.getRange(pinjamRow, 9).setValue("Dikembalikan"); // kolom I = Status

  // Tambah stok di Daftar Buku
  const dataBuku = sheetBuku.getDataRange().getValues();
  for (let i = 1; i < dataBuku.length; i++) {
    if (String(dataBuku[i][0]) === String(idBuku)) {
      sheetBuku.getRange(i + 1, 6).setValue(Number(dataBuku[i][5]) + 1);
      break;
    }
  }

  // ✅ Generate ID Kembali
  const lastRow  = sheetPengembalian.getLastRow();
  const idKembali = "K" + String(lastRow).padStart(4, "0");

  // ✅ appendRow sesuai kolom sheet Pengembalian:
  // ID Kembali | ID Pinjam | Nama Peminjam | Judul Buku | Tanggal Kembali Aktual | Kondisi Buku | Catatan
  sheetPengembalian.appendRow([
    idKembali,       // A - ID Kembali
    idPinjamInput,   // B - ID Pinjam
    namaPeminjam,    // C - Nama Peminjam
    judulBuku,       // D - Judul Buku
    tanggalKembali,  // E - Tanggal Kembali Aktual
    "Baik",          // F - Kondisi Buku (default)
    ""               // G - Catatan (kosong)
  ]);

  // Kirim email konfirmasi
  kirimEmailPengembalian({
    email,
    nama: namaPeminjam,
    idPinjam: idPinjamInput,
    judulBuku,
    tanggalKembali
  });

}

function kirimEmailPeminjaman({ email, nama, idPinjam, judulBuku, tanggal }) {

  const tanggalFormat = Utilities.formatDate(
    tanggal,
    Session.getScriptTimeZone(),
    "dd MMMM yyyy HH:mm"
  );

  const subject = `✅ Konfirmasi Peminjaman Buku - ${judulBuku}`;

  const body = `
    <div style="font-family:Arial,sans-serif;max-width:500px;margin:auto;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;">

      <div style="background:#15803d;padding:20px;text-align:center;">
        <h2 style="color:white;margin:0;">📚 Konfirmasi Peminjaman</h2>
      </div>

      <div style="padding:24px;">
        <p>Halo <strong>${nama}</strong>,</p>
        <p>Peminjaman buku kamu berhasil dicatat. Berikut detailnya:</p>

        <table style="width:100%;border-collapse:collapse;margin:16px 0;">
          <tr style="background:#f3f4f6;">
            <td style="padding:10px;font-weight:bold;width:40%;">ID Peminjaman</td>
            <td style="padding:10px;">${idPinjam}</td>
          </tr>
          <tr>
            <td style="padding:10px;font-weight:bold;">Judul Buku</td>
            <td style="padding:10px;">${judulBuku}</td>
          </tr>
          <tr style="background:#f3f4f6;">
            <td style="padding:10px;font-weight:bold;">Tanggal Pinjam</td>
            <td style="padding:10px;">${tanggalFormat}</td>
          </tr>
        </table>

        <p style="color:#6b7280;font-size:13px;">
          Simpan ID Peminjaman kamu — diperlukan saat pengembalian buku.
        </p>
      </div>

      <div style="background:#f9fafb;padding:16px;text-align:center;font-size:12px;color:#9ca3af;">
        Sistem Perpustakaan
      </div>

    </div>
  `;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: body
  });

}

function kirimEmailPengembalian({ email, nama, idPinjam, judulBuku, tanggalKembali }) {

  const tanggalFormat = Utilities.formatDate(
    tanggalKembali,
    Session.getScriptTimeZone(),
    "dd MMMM yyyy HH:mm"
  );

  const subject = `🔄 Konfirmasi Pengembalian Buku - ${judulBuku}`;

  const body = `
    <div style="font-family:Arial,sans-serif;max-width:500px;margin:auto;border:1px solid #e0e0e0;border-radius:8px;overflow:hidden;">

      <div style="background:#1d4ed8;padding:20px;text-align:center;">
        <h2 style="color:white;margin:0;">🔄 Konfirmasi Pengembalian</h2>
      </div>

      <div style="padding:24px;">
        <p>Halo <strong>${nama}</strong>,</p>
        <p>Pengembalian buku kamu berhasil dicatat. Berikut detailnya:</p>

        <table style="width:100%;border-collapse:collapse;margin:16px 0;">
          <tr style="background:#f3f4f6;">
            <td style="padding:10px;font-weight:bold;width:40%;">ID Peminjaman</td>
            <td style="padding:10px;">${idPinjam}</td>
          </tr>
          <tr>
            <td style="padding:10px;font-weight:bold;">Judul Buku</td>
            <td style="padding:10px;">${judulBuku}</td>
          </tr>
          <tr style="background:#f3f4f6;">
            <td style="padding:10px;font-weight:bold;">Tanggal Kembali</td>
            <td style="padding:10px;">${tanggalFormat}</td>
          </tr>
        </table>

        <p style="color:#16a34a;font-weight:bold;">Terima kasih sudah mengembalikan buku tepat waktu! 🎉</p>
      </div>

      <div style="background:#f9fafb;padding:16px;text-align:center;font-size:12px;color:#9ca3af;">
        Sistem Perpustakaan
      </div>

    </div>
  `;

  MailApp.sendEmail({
    to: email,
    subject: subject,
    htmlBody: body
  });

}

function pinjamBuku(data) {
  const ss = SpreadsheetApp.getActive();
  const sheetBuku  = ss.getSheetByName("Daftar Buku");
  const sheetPinjam = ss.getSheetByName("Peminjaman");

  prosesPeminjaman({
    ss, sheetBuku, sheetPinjam,
    nama: data.nama,
    divisi: data.divisi,
    noHp: data.hp,
    email: data.email,
    idBuku: data.idBuku
  });

  return { success: true, message: "Buku berhasil dipinjam" };
}

function kembalikanBuku(data) {
  const ss = SpreadsheetApp.getActive();

  prosesPengembalian({
    ss,
    sheetBuku: ss.getSheetByName("Daftar Buku"),
    sheetPinjam: ss.getSheetByName("Peminjaman"),
    sheetPengembalian: ss.getSheetByName("Pengembalian"),
    nama: "",
    email: data.email,
    idPinjamInput: data.idPinjam
  });

  return { success: true, message: "Buku berhasil dikembalikan" };
}
