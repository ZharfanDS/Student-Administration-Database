CREATE DATABASE DB_ADMINISTRASI;
GO

USE DB_ADMINISTRASI;
GO

CREATE TABLE Siswa (
  ID INT PRIMARY KEY IDENTITY(1,1), -- ID unik yang otomatis bertambah
  Nama VARCHAR(100) NOT NULL,
  Kelas VARCHAR(20),
  Alamat VARCHAR(255)
);
GO

INSERT INTO Siswa (Nama, Kelas, Alamat) VALUES
('Budi Santoso', '10 IPA 1', 'Jl. Merdeka No. 10'),
('Citra Lestari', '11 IPS 2', 'Jl. Pahlawan No. 5');
GO

SELECT * FROM Siswa;
GO
