# �f�B���N�g�������Z�b�g����
$workingdir = Get-Location
$macrodir = ".\myTest"

# �t�@�C�����i�萔�j���Z�b�g����
$docmName = "$macrodir/CATOVIS_Helper.xlsm"
$files = Get-ChildItem $docmName

# �ŐV�̃t�@�C�����擾����
# �ŐV�̃t�@�C���ionGoing�j���Ȃ��ꍇ�A�������I������
$file = ""
if ($files.Length -lt 1) {
    Write-Host "${macrodir} �Ƀ}�N���L���t�@�C��������܂���"
    $file = $docmName
} else {
    $file = $files[0].FullName
}

Write-Host "�ŏI�t�@�C���F ${file} �ɏ��������s���܂��B��낵���ł����H"
$sw = Read-Host("[y/N]")
if ($sw -ne "y") {
    Write-Host "�L�����Z������܂���"
    Read-host "Type any key:"
    Exit
}

# �}�N�����G�N�X�|�[�g���Ă���
Write-Host "�}�N�����G�N�X�|�[�g���܂�"
cscript.exe .\vbac.wsf decombine /binary $macrodir /source ./code/

.\_commit-macro.ps1