IF OBJECT_ID('dbo.usp_AddLog') IS NOT NULL 
    DROP PROCEDURE [dbo].[usp_AddLog]
GO

CREATE PROCEDURE [dbo].[usp_AddLog]
(
    @_InBox_ReadID          tinyint,                -- 執行 Log 時是使用第幾版
    @_InBox_SPNAME          nvarchar(120),          -- 執行的預存程序名稱
    @_InBox_GroupID         uniqueidentifier,       -- 執行群組代碼
    @_InBox_ExProgram       nvarchar(40),           -- 執行的動作是什麼
    @_InBox_ActionJSON      nvarchar(Max),          -- 執行的過程是什麼
    @_OutBox_ReturnValues   nvarchar(Max) OUTPUT    -- 回傳執行的項目
) 
AS
BEGIN
    -- 宣告預設的變數
    DECLARE @_StoredProgramsNAME nvarchar(100) = 'usp_AddLog';
    DECLARE @_ReturnTable TABLE 
    (
        [RT_Status] bit,                              --執行結果
        [RT_ActionValues] nvarchar(2000)              --回傳結果為何
    );

    BEGIN TRY
        IF(@_InBox_ReadID = 0) 
        BEGIN
            -- 插入執行記錄
            INSERT INTO MyOffice_ExcuteionLog 
            (
                DeLog_StoredPrograms,
                DeLog_GroupID,
                DeLog_ExecutionProgram,
                DeLog_ExecutionInfo
            )
            VALUES
            (
                @_InBox_SPNAME,
                @_InBox_GroupID,
                @_InBox_ExProgram,
                @_InBox_ActionJSON
            );

            -- 查詢回傳 JSON 結果
            SET @_OutBox_ReturnValues =
            (
                SELECT
                    TOP 100 
                    DeLog_AutoID             AS 'AutoID',
                    DeLog_ExecutionProgram   AS 'NAME',
                    DeLog_ExecutionInfo      AS 'Action',
                    DeLog_ExDateTime         AS 'DateTime'
                FROM MyOffice_ExcuteionLog WITH(NOLOCK)
                WHERE DeLog_GroupID = @_InBox_GroupID
                ORDER BY DeLog_AutoID 
                FOR JSON PATH, ROOT('ProgramLog'), INCLUDE_NULL_VALUES
            );
        END;
    END TRY
    BEGIN CATCH
        -- 錯誤處理，這裡可以記錄錯誤訊息或者回傳錯誤訊息
        SET @_OutBox_ReturnValues = 'Error occurred: ' + ERROR_MESSAGE();
    END CATCH
END;