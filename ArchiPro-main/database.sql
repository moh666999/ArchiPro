BEGIN TRANSACTION;
CREATE TABLE archive(

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            file_number TEXT UNIQUE,

            subject TEXT,

            department TEXT,

            archive_type TEXT,

            document_date TEXT,

            file_path TEXT,

            created_at TEXT,

            updated_at TEXT,

            status TEXT

        , file_name TEXT, storage_location TEXT, document_file_path TEXT, document_type TEXT);
INSERT INTO "archive" VALUES(5,'001','ANAAAAAAAAAAAAA','','الكتروني','','HNAAK','2026-09-16 04:24:27','2026-09-16 04:24:27','Active','ANA',NULL,'Archive_Files\001_Math112_Final_Revision_2022.pdf','سري');
INSERT INTO "archive" VALUES(6,'002','.........','','الكتروني','','...','2026-09-16 14:18:30','2026-09-16 14:18:30','Active','test',NULL,'Archive_Files\002_001_Math112_Final_Revision_2022.pdf','عادي');
CREATE TABLE history(

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            file_number TEXT,

            action TEXT,

            action_date TEXT

        , username
                TEXT, device_name
                TEXT);
INSERT INTO "history" VALUES(22,'001','Added File','2026-09-16 04:24:27','emaadtalaat321','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0');
INSERT INTO "history" VALUES(23,'001','Viewed File','2026-09-16 04:24:34','emaadtalaat321','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0');
INSERT INTO "history" VALUES(24,'001','Viewed File','2026-09-16 04:24:42','emaadtalaat321','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0');
INSERT INTO "history" VALUES(25,'001','Viewed File','2026-09-16 14:10:49','emaadtalaat321','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0');
INSERT INTO "history" VALUES(26,'001','Viewed File','2026-09-16 14:10:55','emaadtalaat321','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0');
INSERT INTO "history" VALUES(27,'001','Downloaded File','2026-09-16 14:11:01','emaadtalaat321','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0');
INSERT INTO "history" VALUES(28,'002','Added File','2026-09-16 14:18:31','emaadtalaat321','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0');
INSERT INTO "history" VALUES(29,'002','Viewed File','2026-09-16 14:18:50','emaadtalaat321','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0');
INSERT INTO "history" VALUES(30,'002','Viewed File','2026-09-16 14:18:54','emaadtalaat321','Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36 Edg/153.0.0.0');
CREATE TABLE users(

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            username TEXT UNIQUE,
            password TEXT

        , role
                TEXT DEFAULT 'normal_user', can_view
                INTEGER DEFAULT 1, can_download
                INTEGER DEFAULT 0, can_add
                INTEGER DEFAULT 0, can_edit
                INTEGER DEFAULT 0, can_delete
                INTEGER DEFAULT 0, can_normal INTEGER DEFAULT 0, can_secret INTEGER DEFAULT 0, can_top_secret INTEGER DEFAULT 0);
INSERT INTO "users" VALUES(46,'emaadtalaat321','ETN2110','main_user',1,1,1,1,1,0,0,0);
INSERT INTO "users" VALUES(148,'Perryy','12345','normal_user',1,0,0,0,1,0,0,0);
DELETE FROM "sqlite_sequence";
INSERT INTO "sqlite_sequence" VALUES('archive',6);
INSERT INTO "sqlite_sequence" VALUES('users',148);
INSERT INTO "sqlite_sequence" VALUES('history',30);
COMMIT;