SELECT DISTINCT 
 '../../reports/' || :METERID || '?sDate=' ||  strftime('%Y-%m-%d', datetime(measdate,'start of month'))  || '&eDate=' || strftime('%Y-%m',datetime(datetime(measdate,'start of month'),'+1 months'))  || '-01' 
 AS REPORT_URL, 
 strftime('%Y-%m', measdate) AS REPORT_NAME
FROM READINGS
WHERE METER_ID = :METERID
