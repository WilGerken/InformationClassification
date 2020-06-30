select count(*) as DELETE_CNT from [prod].FS_SERVER_OBJ_FLAG_VEW where FLAG_STATUS_TXT = 'Executed'

select * from [prod].FS_SERVER_OBJ_RESULT_VEW order by SERVER_URL, SERVER_ID, OBJ_TYPE_TXT, FLAG_STATUS_TXT

