IF OBJECT_ID('dbo.usp_AddLog') IS NOT NULL 
    DROP PROCEDURE [dbo].[usp_AddLog]
GO

CREATE PROCEDURE [dbo].[usp_AddLog]
(
    @_InBox_ReadID          tinyint,                -- ���� Log �ɬO�ϥβĴX��
    @_InBox_SPNAME          nvarchar(120),          -- ���檺�w�s�{�ǦW��
    @_InBox_GroupID         uniqueidentifier,       -- ����s�եN�X
    @_InBox_ExProgram       nvarchar(40),           -- ���檺�ʧ@�O����
    @_InBox_ActionJSON      nvarchar(Max),          -- ���檺�L�{�O����
    @_OutBox_ReturnValues   nvarchar(Max) OUTPUT    -- �^�ǰ��檺����
) 
AS
BEGIN
    -- �ŧi�w�]���ܼ�
    DECLARE @_StoredProgramsNAME nvarchar(100) = 'usp_AddLog';
    DECLARE @_ReturnTable TABLE 
    (
        [RT_Status] bit,                              --���浲�G
        [RT_ActionValues] nvarchar(2000)              --�^�ǵ��G����
    );

    BEGIN TRY
        IF(@_InBox_ReadID = 0) 
        BEGIN
            -- ���J����O��
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

            -- �d�ߦ^�� JSON ���G
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
        -- ���~�B�z�A�o�̥i�H�O�����~�T���Ϊ̦^�ǿ��~�T��
        SET @_OutBox_ReturnValues = 'Error occurred: ' + ERROR_MESSAGE();
    END CATCH
END;