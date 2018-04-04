robocopy \\fpserver\e$\ACCT E:\ACCT /COPYALL /r:1 /w:1 /V /e
robocopy \\fpserver\e$\data E:\Data /COPYALL /r:1 /w:1 /V /e
robocopy \\fpserver\e$\KThomas E:\KThomas /COPYALL /r:1 /w:1 /V /e
robocopy "\\fpserver\e$\Manul Backups" E:\Manual Backups /COPYALL /r:1 /w:1 /V /e
robocopy \\fpserver\e$\SQLBackups E:\SQLBackups /COPYALL /r:1 /w:1 /V /e
robocopy \\fpserver\e$\SQLMaint E:\SQLMaint /COPYALL /r:1 /w:1 /V /e
Pause