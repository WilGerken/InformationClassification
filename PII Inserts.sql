INSERT INTO dbo.InfoClass
select 'PIIFiles' AS filename,
GETDATE() AS importdate,
action,
filefolder,
'helzl0',
GETDATE(),
'1.0.0.0'
from (
select distinct action, newdrive as filefolder from 
(SELECT ACTION,
       CASE
           WHEN UPPER(LEFT([FilePath], 27)) = UPPER('\\ramancss003.bp.com\group\')
           THEN 'S:\' + SUBSTRING([FilePath], 28, LEN([FilePath]) - 27)
           WHEN UPPER(LEFT([FilePath], 20)) = UPPER('\\ramancss003\group\')
           THEN 'S:\' + SUBSTRING([FilePath], 21, LEN([FilePath]) - 20)
           WHEN UPPER(LEFT([FilePath], 27)) = UPPER('\\xt3ancss206.bp.com\group\')
           THEN 'T:\' + SUBSTRING([FilePath], 28, LEN([FilePath]) - 27)
           WHEN UPPER(LEFT([FilePath], 20)) = UPPER('\\xt3ancss206\group\')
           THEN 'T:\' + SUBSTRING([FilePath], 21, LEN([FilePath]) - 20)
           WHEN UPPER(LEFT([FilePath], 21)) = UPPER('\\ramancss003\GroupU\')
           THEN 'U:\' + SUBSTRING([FilePath], 22, LEN([FilePath]) - 21)
           WHEN UPPER(LEFT([FilePath], 21)) = UPPER('\\rampbess003\GroupU\')
           THEN 'U:\' + SUBSTRING([FilePath], 22, LEN([FilePath]) - 21)
           WHEN UPPER(LEFT([FilePath], 28)) = UPPER('\\ramancss004.bp.com\groupV\')
           THEN 'V:\' + SUBSTRING([FilePath], 29, LEN([FilePath]) - 28)
           WHEN UPPER(LEFT([FilePath], 28)) = UPPER('\\rampmess004.bp.com\groupV\')
           THEN 'V:\' + SUBSTRING([FilePath], 29, LEN([FilePath]) - 28)
           WHEN UPPER(LEFT([FilePath], 21)) = UPPER('\\rampbess004\groupV\')
           THEN 'V:\' + SUBSTRING([FilePath], 22, LEN([FilePath]) - 21)
           WHEN UPPER(LEFT([FilePath], 26)) = UPPER('\\xt3ancss206.bp.com\MAOG\')
           THEN 'M:\' + SUBSTRING([FilePath], 27, LEN([FilePath]) - 26)
           WHEN UPPER(LEFT([FilePath], 19)) = UPPER('\\xt3ancss206\MAOG\')
           THEN 'M:\' + SUBSTRING([FilePath], 20, LEN([FilePath]) - 19)
           ELSE '?'
       END AS newdrive, 
       filepath
FROM [InformationClassification].[dbo].[pii-resume-cv]
union
SELECT ACTION,
       CASE
           WHEN UPPER(LEFT([FilePath], 27)) = UPPER('\\ramancss003.bp.com\group\')
           THEN 'S:\' + SUBSTRING([FilePath], 28, LEN([FilePath]) - 27)
           WHEN UPPER(LEFT([FilePath], 20)) = UPPER('\\ramancss003\group\')
           THEN 'S:\' + SUBSTRING([FilePath], 21, LEN([FilePath]) - 20)
           WHEN UPPER(LEFT([FilePath], 27)) = UPPER('\\xt3ancss206.bp.com\group\')
           THEN 'T:\' + SUBSTRING([FilePath], 28, LEN([FilePath]) - 27)
           WHEN UPPER(LEFT([FilePath], 20)) = UPPER('\\xt3ancss206\group\')
           THEN 'T:\' + SUBSTRING([FilePath], 21, LEN([FilePath]) - 20)
           WHEN UPPER(LEFT([FilePath], 21)) = UPPER('\\ramancss003\GroupU\')
           THEN 'U:\' + SUBSTRING([FilePath], 22, LEN([FilePath]) - 21)
           WHEN UPPER(LEFT([FilePath], 21)) = UPPER('\\rampbess003\GroupU\')
           THEN 'U:\' + SUBSTRING([FilePath], 22, LEN([FilePath]) - 21)
           WHEN UPPER(LEFT([FilePath], 28)) = UPPER('\\ramancss004.bp.com\groupV\')
           THEN 'V:\' + SUBSTRING([FilePath], 29, LEN([FilePath]) - 28)
           WHEN UPPER(LEFT([FilePath], 28)) = UPPER('\\rampmess004.bp.com\groupV\')
           THEN 'V:\' + SUBSTRING([FilePath], 29, LEN([FilePath]) - 28)
           WHEN UPPER(LEFT([FilePath], 21)) = UPPER('\\rampbess004\groupV\')
           THEN 'V:\' + SUBSTRING([FilePath], 22, LEN([FilePath]) - 21)
           WHEN UPPER(LEFT([FilePath], 26)) = UPPER('\\xt3ancss206.bp.com\MAOG\')
           THEN 'M:\' + SUBSTRING([FilePath], 27, LEN([FilePath]) - 26)
           WHEN UPPER(LEFT([FilePath], 19)) = UPPER('\\xt3ancss206\MAOG\')
           THEN 'M:\' + SUBSTRING([FilePath], 20, LEN([FilePath]) - 19)
           ELSE '?'
       END AS newdrive, 
       filepath
FROM [InformationClassification].[dbo].[PIIResultsLinkedtoEVE_EG-202006]
union
SELECT ACTION,
       CASE
           WHEN UPPER(LEFT([FilePath], 27)) = UPPER('\\ramancss003.bp.com\group\')
           THEN 'S:\' + SUBSTRING([FilePath], 28, LEN([FilePath]) - 27)
           WHEN UPPER(LEFT([FilePath], 20)) = UPPER('\\ramancss003\group\')
           THEN 'S:\' + SUBSTRING([FilePath], 21, LEN([FilePath]) - 20)
           WHEN UPPER(LEFT([FilePath], 27)) = UPPER('\\xt3ancss206.bp.com\group\')
           THEN 'T:\' + SUBSTRING([FilePath], 28, LEN([FilePath]) - 27)
           WHEN UPPER(LEFT([FilePath], 20)) = UPPER('\\xt3ancss206\group\')
           THEN 'T:\' + SUBSTRING([FilePath], 21, LEN([FilePath]) - 20)
           WHEN UPPER(LEFT([FilePath], 21)) = UPPER('\\ramancss003\GroupU\')
           THEN 'U:\' + SUBSTRING([FilePath], 22, LEN([FilePath]) - 21)
           WHEN UPPER(LEFT([FilePath], 21)) = UPPER('\\rampbess003\GroupU\')
           THEN 'U:\' + SUBSTRING([FilePath], 22, LEN([FilePath]) - 21)
           WHEN UPPER(LEFT([FilePath], 28)) = UPPER('\\ramancss004.bp.com\groupV\')
           THEN 'V:\' + SUBSTRING([FilePath], 29, LEN([FilePath]) - 28)
           WHEN UPPER(LEFT([FilePath], 28)) = UPPER('\\rampmess004.bp.com\groupV\')
           THEN 'V:\' + SUBSTRING([FilePath], 29, LEN([FilePath]) - 28)
           WHEN UPPER(LEFT([FilePath], 21)) = UPPER('\\rampbess004\groupV\')
           THEN 'V:\' + SUBSTRING([FilePath], 22, LEN([FilePath]) - 21)
           WHEN UPPER(LEFT([FilePath], 26)) = UPPER('\\xt3ancss206.bp.com\MAOG\')
           THEN 'M:\' + SUBSTRING([FilePath], 27, LEN([FilePath]) - 26)
           WHEN UPPER(LEFT([FilePath], 19)) = UPPER('\\xt3ancss206\MAOG\')
           THEN 'M:\' + SUBSTRING([FilePath], 20, LEN([FilePath]) - 19)
           ELSE '?'
       END AS newdrive, 
       filepath
FROM [InformationClassification].[dbo].[Sheet1]
union

SELECT ACTION,
       CASE
           WHEN UPPER(LEFT([FilePath], 27)) = UPPER('\\ramancss003.bp.com\group\')
           THEN 'S:\' + SUBSTRING([FilePath], 28, LEN([FilePath]) - 27)
           WHEN UPPER(LEFT([FilePath], 20)) = UPPER('\\ramancss003\group\')
           THEN 'S:\' + SUBSTRING([FilePath], 21, LEN([FilePath]) - 20)
           WHEN UPPER(LEFT([FilePath], 27)) = UPPER('\\xt3ancss206.bp.com\group\')
           THEN 'T:\' + SUBSTRING([FilePath], 28, LEN([FilePath]) - 27)
           WHEN UPPER(LEFT([FilePath], 20)) = UPPER('\\xt3ancss206\group\')
           THEN 'T:\' + SUBSTRING([FilePath], 21, LEN([FilePath]) - 20)
           WHEN UPPER(LEFT([FilePath], 21)) = UPPER('\\ramancss003\GroupU\')
           THEN 'U:\' + SUBSTRING([FilePath], 22, LEN([FilePath]) - 21)
           WHEN UPPER(LEFT([FilePath], 21)) = UPPER('\\rampbess003\GroupU\')
           THEN 'U:\' + SUBSTRING([FilePath], 22, LEN([FilePath]) - 21)
           WHEN UPPER(LEFT([FilePath], 28)) = UPPER('\\ramancss004.bp.com\groupV\')
           THEN 'V:\' + SUBSTRING([FilePath], 29, LEN([FilePath]) - 28)
           WHEN UPPER(LEFT([FilePath], 28)) = UPPER('\\rampmess004.bp.com\groupV\')
           THEN 'V:\' + SUBSTRING([FilePath], 29, LEN([FilePath]) - 28)
           WHEN UPPER(LEFT([FilePath], 21)) = UPPER('\\rampbess004\groupV\')
           THEN 'V:\' + SUBSTRING([FilePath], 22, LEN([FilePath]) - 21)
           WHEN UPPER(LEFT([FilePath], 26)) = UPPER('\\xt3ancss206.bp.com\MAOG\')
           THEN 'M:\' + SUBSTRING([FilePath], 27, LEN([FilePath]) - 26)
           WHEN UPPER(LEFT([FilePath], 19)) = UPPER('\\xt3ancss206\MAOG\')
           THEN 'M:\' + SUBSTRING([FilePath], 20, LEN([FilePath]) - 19)
           ELSE '?'
       END AS newdrive, 
       filepath
FROM [InformationClassification].[dbo].[tbl_PST_Master_List]) d) dd