echo off

Rem ���Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ���
Rem �E���s�R�[�h���uLF�v�݂̂ɂȂ��Ă���Ɛ��������삵�Ȃ����߁A
Rem   ���s�R�[�h���uCR+LF�v�ɂ��邱��
Rem   github����_�E�����[�h���Ă����LF�ɂȂ��Ă��܂��͗l
Rem ���Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ��Ӓ���


Rem ----------------------- �ݒ� -----------------------
Rem �R���p�C���Ώۃ\�[�X�t�H���_�ݒ� 
Rem ���R���p�C���Ώۃ\�[�X�̃f�t�H���g�̐ݒ�ł�
Set CompileSourceFolder=%~dp0
Rem sln�t�@�C����                    
Rem ���\�����[�V�����t�@�C�����̃f�t�H���g�̐ݒ�ł�
Set SlnFileName=ImageForClipboard.sln
Rem .NetFramework Version�w��        
Rem ��.NetFramework�̃o�[�W�����w��̃f�t�H���g�ݒ�ł� �u%windir%\Microsoft.NET\Framework\�v�z���̃t�H���_�����w�肵�ĉ�����
Set FrameworkVersion=v4.0.30319
Rem ----------------------- �ݒ� -----------------------

Rem �^�C�g���\��
:DisplayTitle

    echo;
    echo  ********************************************************
    echo    ImageForClipboard�R���p�C���p�o�b�`
    echo      ImageForClipboard�̃��r���h���s��exe�t�@�C�����쐬���܂�
    echo *********************************************************

Rem �t�H���_�E�t�@�C���w�菈��
:SpecifyCompileFolderAndSlnFile

    echo;
    echo  ********************************************************
    echo  �R���p�C���t�H���_�E�\�����[�V�����t�@�C���w�菈��
    echo  ********************************************************

    echo;
    echo ���R���p�C���Ώۃ\�[�X�t�H���_���w�肵�Ă�������
    echo �������w�肵�Ȃ��ꍇ�́F�u%CompileSourceFolder%�v�t�H���_�ɂȂ�܂�
    echo;
    Set /p CompileSourceFolder="�t�H���_�̓��́@���@"

    echo;
    echo ���\�����[�V�����t�@�C�������w�肵�Ă�������
    echo �������w�肵�Ȃ��ꍇ�́F�u%SlnFileName%�v�t�@�C���ɂȂ�܂�
    echo;
    Set /p SlnFileName="�t�@�C�����̓��́@���@"

Rem �ݒ�̊m�F
:DisplayConfiguration

    echo;
    echo  ********************************************************
    echo  �R���p�C���t�H���_�@�@�@�F%CompileSourceFolder%
    echo  �\�����[�V�����t�@�C�����F%SlnFileName%
    echo  ********************************************************
    echo;
    
    echo �����L���b�Z�[�W��(�uY�v���́uy�v)�ȊO�̓L�����Z������܂�
    Set /p RunContinueResult="��L�����g�p���ď��������s���܂����H(y/n)�@���@"

    Rem �啶��/�������ϊ�(Y�ȊO�͑S�ăL�����Z������) 
    Set RunContinueResult=%RunContinueResult:y=Y%%
    
    Rem Y�ȊO�̓��͂̎��̓R���p�C���t�H���_�E�\�����[�V�����t�@�C���w�菈����
    If /i Not %RunContinueResult%==Y Goto SpecifyCompileFolderAndSlnFile

Rem �\�[�X�̃R���p�C������
:CompileProcess

    echo;
    echo  ********************************************************
    echo  �R���p�C������
    echo  ********************************************************
    
    echo;
    echo  ���J�����g�t�H���_��ύX���܂��c�c
    echo;
    cd %CompileSourceFolder%

    echo;
    echo  ���R���p�C�������s���܂��c�c
    echo;
    %windir%\Microsoft.NET\Framework\%FrameworkVersion%\MSBuild.exe %SlnFileName% /t:Rebuild /p:Configuration=Release

    echo;
    echo  ���������������܂����I�I
    echo;

    pause

Rem �I������
:EndProcess

    Rem�����̏I��
    exit