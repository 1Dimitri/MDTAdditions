:: $Id$
:: $Rev$
:: $URL$
:: $Date$
set p=HKLM\Software\MDTX2010\SMSEnv
reg add %p% /t reg_sz /v UserID /d "%UserID%" /f
reg add %p% /t reg_sz /v UserID /d "%UserDomain%" /f
