-- ================================================
-- Template generated from Template Explorer using:
-- Create Procedure (New Menu).SQL
--
-- Use the Specify Values for Template Parameters 
-- command (Ctrl-Shift-M) to fill in the parameter 
-- values below.
--
-- This block of comments will not be included in
-- the definition of the procedure.
-- ================================================
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
-- =============================================
-- Author:		<Author,,Name>
-- Create date: <Create Date,,>
-- Description:	<Description,,>
-- =============================================
CREATE PROCEDURE SQL_Table_To_VB_Class
	-- Add the parameters for the stored procedure here
	 @TableName sysname
AS
BEGIN
	-- SET NOCOUNT ON added to prevent extra result sets from
	-- interfering with SELECT statements.
	SET NOCOUNT ON;

    -- Insert statements for procedure here
declare @result varchar(max) = 'Public Class ' + @TableName 
select @Result = @Result +' '+ primaryKey+'
Property '+ ColumnName+' as '+ RTRIM(Columntype) + NullableSign+' '
from
(
    select 
        replace(col.name, ' ', '_') ColumnName,
        column_id,
        case typ.name 
             when 'bigint' then 'long'
        when 'binary' then 'byte'
        when 'bit' then 'boolean'
        when 'char' then 'string'
        when 'date' then 'DateTime'
        when 'datetime' then 'DateTime'
        when 'datetime2' then 'DateTime'
        when 'datetimeoffset' then 'DateTimeOffset'
        when 'decimal' then 'decimal'
        when 'float' then 'double'
        when 'image' then 'byte'
        when 'int' then 'integer'
        when 'money' then 'decimal'
        when 'nchar' then 'string'
        when 'ntext' then 'string'
        when 'numeric' then 'decimal'
        when 'nvarchar' then 'string'
        when 'real' then 'double'
        when 'smalldatetime' then 'DateTime'
        when 'smallint' then 'short'
        when 'smallmoney' then 'decimal'
        when 'text' then 'string'
        when 'time' then 'TimeSpan'
        when 'timestamp' then 'DateTime'
        when 'tinyint' then 'byte'
        when 'uniqueidentifier' then 'Guid'
        when 'varbinary' then 'byte'
        when 'varchar' then 'string'
        else 'UNKNOWN_' + typ.name
        END + CASE WHEN col.is_nullable=1 AND typ.name NOT IN ('binary', 'varbinary', 'image', 'text', 'ntext', 'varchar', 'nvarchar', 'char', 'nchar') THEN ' ' ELSE '' END ColumnType,
        			case 
		when col.is_nullable = 1 and typ.name in ('bigint', 'bit', 'date', 'datetime', 'datetime2', 'datetimeoffset', 'decimal', 'float', 'int', 'money', 'numeric', 'real', 'smalldatetime', 'smallint', 'smallmoney', 'time', 'tinyint', 'uniqueidentifier') 
		then '?' 
		else '' 
	end NullableSign,
				    case
                when col.name in (SELECT column_name
                                    FROM INFORMATION_SCHEMA.TABLE_CONSTRAINTS AS TC
                                    inner join 
                                        INFORMATION_SCHEMA.KEY_COLUMN_USAGE AS KU
                                            ON TC.CONSTRAINT_TYPE = 'PRIMARY KEY' AND
                                                TC.CONSTRAINT_NAME = KU.CONSTRAINT_NAME and 
                                                KU.table_name=@TableName)
                                            then '
            <PrimaryKey>'
                else ' '	
                end primaryKey,
		colDesc.colDesc AS ColumnDesc
    from sys.columns col
        join sys.types typ on
            col.system_type_id = typ.system_type_id AND col.user_type_id = typ.user_type_id
    OUTER APPLY (
    SELECT TOP 1 CAST(value AS NVARCHAR(max)) AS colDesc
    FROM
       sys.extended_properties
    WHERE
       major_id = col.object_id
       AND
       minor_id = COLUMNPROPERTY(major_id, col.name, 'ColumnId')
    ) colDesc            
    where object_id = object_id(@TableName)
) t
order by column_id

set @result = @result  + '
End Class'

print @result
END
GO
