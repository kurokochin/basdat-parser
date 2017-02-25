Parser dari .xlsx untuk mata kuliah basis data


Penggunaan:
- git clone https://github.com/kurokochin/basdat-parser.git
- pindahkan file xlsx ke folder basdat-parse (linux : mv data_sirenkom.xlsx basdat-parser)
- [python3/python] parser.py [nama_file_xlsx] [sheet1/sheet2/sheet3/sheet4/..] [nama_file_output] [nama_table] 
Contoh: python parser.py data_sirenkom.xlsx sheet1 petugas.txt PETUGAS

Hasil:
petugas.txt

Isinya:  
INSERT INTO PETUGAS (id_petugas, nama_lengkap, alamat, no_telp) VALUES ( 'P00001' ,'Samsudin' ,'Jl. Samudera I No.7, Depok' ,'0218313541' );  
INSERT INTO PETUGAS (id_petugas, nama_lengkap, alamat, no_telp) VALUES ( 'P00002' ,'Siska Wijaya' ,'Jl. Pinus Raya No.21, Depok' ,'0213456234' );  
INSERT INTO PETUGAS (id_petugas, nama_lengkap, alamat, no_telp) VALUES ( 'P00003' ,'Rosiana' ,'Jl. Gardu Timur No.89, Bekasi' ,'0219876345' );  
INSERT INTO PETUGAS (id_petugas, nama_lengkap, alamat, no_telp) VALUES ( 'P00004' ,'Chitra Ayu' ,'Jl. Pinus Raya No.10, Depok' ,'0218356742' );  
INSERT INTO PETUGAS (id_petugas, nama_lengkap, alamat, no_telp) VALUES ( 'P00005' ,'Asywendy Putri' ,'Jl. Pelataran No. 17, Depok' ,'0218886241' );  
INSERT INTO PETUGAS (id_petugas, nama_lengkap, alamat, no_telp) VALUES ( 'P00006' ,'Ferdiansyah' ,'Jl. Tegar Beriman No. 38, Bekasi' ,'0213456197' );  
INSERT INTO PETUGAS (id_petugas, nama_lengkap, alamat, no_telp) VALUES ( 'P00007' ,'Ardinata Putra' ,'Jl. Pegangsaan Timur No. 1, Jakarta' ,'0219645789' );  
INSERT INTO PETUGAS (id_petugas, nama_lengkap, alamat, no_telp) VALUES ( 'P00008' ,'Aryan Sapta' ,'Jl. Gelatik Raya No. 42, Jakarta' ,'0213946759' );  
INSERT INTO PETUGAS (id_petugas, nama_lengkap, alamat, no_telp) VALUES ( 'P00009' ,'Erwin' ,'Jl. Pegangsaan Timur No. 22, Jakarta' ,'0213342789' );  
INSERT INTO PETUGAS (id_petugas, nama_lengkap, alamat, no_telp) VALUES ( 'P00010' ,'Laila Sari' ,'Jl. Juragan Sinda II No 24, Depok' ,'0214358575' );  
