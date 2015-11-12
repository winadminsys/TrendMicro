SELECT
       DB_NAME(vfs.database_id) AS DB
,      Latency = AVG(CASE WHEN (num_of_reads = 0 AND num_of_writes = 0) 
                    THEN 0 ELSE (io_stall / (num_of_reads + num_of_writes)) END)
,      ReadLatency = AVG(CASE WHEN num_of_reads = 0
                    THEN 0 ELSE (io_stall_read_ms / num_of_reads) END)
,      WriteLatency = AVG(CASE WHEN num_of_writes = 0
                    THEN 0 ELSE (io_stall_write_ms / num_of_writes) END)
,      AvgBytesPerRead = AVG(CASE WHEN num_of_reads = 0
                    THEN 0 ELSE (num_of_bytes_read / num_of_reads) END)
,      AvgBytesPerWrite = AVG(CASE WHEN num_of_writes = 0
                    THEN 0 ELSE (num_of_bytes_written / num_of_writes) END)
FROM sys.dm_io_virtual_file_stats(NULL, NULL) AS vfs
       JOIN sys.master_files mf ON vfs.database_id = mf.database_id AND vfs.file_id = mf.file_id
where type = 0
group by DB_NAME(vfs.database_id)
ORDER BY DB
