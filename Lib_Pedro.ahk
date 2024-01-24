/************************************************************************
 * @description Biblioteca com funções auxiliares ao script Pedro.ahk
 * @author Pedro Henrique C. Xavier
 * @date 2024/01/23
 * @version 2.0.11
 ***********************************************************************/

AbrirClipBoard(*) => Notepad2('/c')
AjusteCfopCst(*) => Run('\Conferir_arquivo_XML\Ajustar XML CFOP e CST inconsistentes.ahk')
AlwaysOnTop(*) => WinSetAlwaysOnTop(-1, 'A')
BlocoKZerado(*) => Run('Conferir_SPED_Fiscal\Inserir Bloco K zerado.ahk')
BuscaCNPJ(*) => Run('Busca_CNPJ [fiscal].ahk')
CalcClipBoard(*) => ClipTools(9)
CalcDifal(*) => RunAct('Calculadora DIFAL.ahk', 'DIFAL Simplificado', 1)
CalcSNacional(*) => RunAct('Calculadora Simples Nacional.ahk', 'Calc S Nacional', 1)
CalcST(*) => RunAct('Calculadora ST.ahk', 'Calculadora ST simplificada', 1)
CapturarClipBoard(*) => Notepad2('/b')
ConfRecSpedFiscal(*) => Run('Conferir_SPED_Fiscal\Conferir Recibos SPED Fiscal.ahk')
ConfSpedEsp(*) => Run('Conferir_SPED_Fiscal\SPED x Espelho NG [Entradas].ahk')
ConfSpedFiscal(*) => Run('Conferir_SPED_Fiscal\Inconsistências SPED Fiscal.ahk')
ConferirEspelho(*) => RunAct('Conferir_Espelho_Lançamento\Conferir Espelho Lançamento.ahk', 'Opções de validação', 1)
ConsultaCFOP(*) => RunAct('ConsultaCFOP [Rich].ahk', 'Consulta CFOP', 1)
ConverterPDF(*) => Run('Converter PDF em Texto.ahk')
DigClip(*) => (SetKeyDelay(20, 20), SendEvent(A_Clipboard))
DocObgFiscais(*) => Run('..\Depto Fiscal\Obrigações Fiscais.xlsx')
DocRelSimples(*) => Run('\\srvrg-saas\rede\RELACAO EMPRESAS - SIMPLES NACIONAL\RELACAO EMPRESAS SIMPLES NACIONAL.xlsx')
EditScript(*) => Edit()
EnvAcessorias(*) => Run('https://app.acessorias.com/sysmain.php?m=144&act=e&i=2')
EveryCert(*) => ControlSetEverything('*.pfx dm:>=' A_YYYY A_MM ' ')
EveryDocDig(*) => ControlSetEverything('folder: parent:"\\srvrg-saas\rede\DIGITALIZADOS - DOCUMENTOS DIVERSOS\" ')
EveryNotas(*) => ControlSetEverything('file:"\\srvrg-saas\rede\PEDRO\Depto Fiscal\Notas\" ')
EveryPedro(*) => ControlSetEverything('"\\srvrg-saas\rede\Pedro" ')
EveryXML(*) => ControlSetEverything('folder: parent:"\\srvrg-saas\rede\NF ELETRONICA - ARQUIVO XML\" ')
ExtrairXML(*) => Run('Extrair_XML_ZIP.ahk')
ExtrairXmlZip(*) => Run('Extrair_XML_ZIP.ahk')
ExtrairZip(*) => Run('ExtrairArquivoCompactado.ahk')
FerramentasTexto(*) => RunAct('Ferramentas de texto.ahk', 'Manipular Strings', 1)
FoldArqXML(*) => openfolder('\\srvrg-saas\rede\NF ELETRONICA - ARQUIVO XML')
FoldDocDig(*) => openfolder('\\srvrg-saas\rede\DIGITALIZADOS - DOCUMENTOS DIVERSOS')
FoldDocPedro(*) => openfolder('\\srvrg-saas\rede\PEDRO')
FoldDownload(*) => openfolder(A_Desktop '\..\Downloads')
FoldExtratos(*) => Run('\\srvrg-saas\rede\Extratos e Boletos Bancários')
FoldImpostos(*) => Run('\\srvrg-saas\rede\PEDRO\Impostos ' mesPassado('MMMM yyyy'))
FoldRecSintegras(*) => Run('\\srvrg-saas\rede\Recibos Sintegra')
FoldScan(*) => openfolder('\\srvrg-saas\rede\Scan')
FoldScripts(*) => Run('\\srvrg-saas\rede\PEDRO\Scripts')
FoldSintegrasClientes(*) => Run('\\srvrg-saas\rede\SINTEGRA - CLIENTES')
FoldXML(*) => ('\\srvrg-saas\rede\PEDRO\XML')
HelpAHK(*) => Run('C:\Program Files\AutoHotkey\v2\AutoHotkey.chm')
ImpArquivosPDF(*) => Run(Format('{} {}\ImprimirArquivosPDF.ahk "{}"', A_AhkPath, A_WorkingDir, LastHwnd()))
IncoSped(*) => Run('Conferir_SPED_Fiscal\Inconsistências Sped Fiscal.ahk')

;InlineCalc(*) =>  WinExist('Calculadora Inline ahk_class AutoHotkeyGUI') ?  WinClose() :  Run('Calculadora Inline.ahk')
InsAspasDuplas(*) => ClipTools(2)
InsAspasSimples(*) => ClipTools(1)
InsClipBoard(*) => (CoordMode('Mouse'), MouseGetPos(&X, &Y), InputGui('X' X ' Y' Y))
InsMaiusculas(*) => ClipTools(4)
InsMinusculas(*) => ClipTools(5)
InsParenteses(*) => ClipTools(3)
InsTitulo(*) => ClipTools(6)
LastHwnd(*) => GetExplorerPath(WinExist('A'))
ListarArqPasta(*) => Run(Format('{} {}\ListarArquivosdePasta.ahk "{}"', A_AhkPath, A_WorkingDir, LastHwnd()))
LoginARAdm(*) => SendEvent('Administrador{Tab}Tech123/qwe@{Tab 6}{Enter}')
LoginARMateus(*) => SendEvent('mateus.duarte{Tab}@mat123{Tab 6}{Enter}')
LoginARPed(*) => SendEvent('@ped123{Tab 3}{Enter}')
InsModoFrase(*) => ClipTools(7)
MoverXML(*) => Run('Conferir_arquivo_XML\MoverXML.ahk')
MoverXMLSubPedro(*) => Run('Mover XML subpasta Pedro.ahk')
Mute(*) => SoundSetMute(-1)
Notepad2(flag) => Run(A_AppData '\Notepad2\Notepad2.exe ' flag)
OpenScript(*) => ListVars()
OpenSiare(*) => Run('"C:\Program Files\Google\Chrome\Application\chrome.exe" https://www2.fazenda.mg.gov.br/sol/')
OpenSintegra(*) => Run('https://dfe-portal.svrs.rs.gov.br/NFE/CCC')
PauseScript(*) => Pause(-1)
ProtFiscal(*) => Run('Protocolos\Protocolo Fiscal.ahk')
QuitApp(*) => Send('!{F4}')
ReloadScript(*) => Reload()
RemAspasParent(*) => ClipTools(8)
RemoverPontuacao(*) => SendEvent(RegExReplace(A_Clipboard, '(*UCP)([^\w,]|R\$)'))
RemoverQuebrasLinha(*) => SendEvent(RegExReplace(A_Clipboard, '[\n\r\t]|^\s*|\s+$|\s{2,}'))
RenChave(*) => Run(Format('{} {}\RenomearChaveAcesso.ahk "{}"', A_AhkPath, A_WorkingDir, LastHwnd()))
RenPDF(*) => Run(Format('{} {}\Lib\RenomearPDF.ahk "{}"', A_AhkPath, A_WorkingDir, LastHwnd()))
RunAgenda(*) => RunAct('ContatosGoogle.ahk', 'Agenda de Telefones', 1)
RunEdge(*) => Run('C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe')
RunGmail(*) => Run('https://mail.google.com/mail/u/0/#inbox')
RunLembretes(*) => RunAct('data\lembretes.ini')
RunNotepad(*) => RunAct('notepad.exe')
RunSkype(*) => RunAct('skype')
SairScript(*) => ExitApp()
SiegHub(*) => Run('https://hub.sieg.com/')
SiegOpen(*) => Run('https://cofre.sieg.com/')
SimNacional(*) => Run('https://www8.receita.fazenda.gov.br/SimplesNacional/controleAcesso/Autentica.aspx?id=60')
SpedvsEspelho(*) => Run('Conferir_SPED_Fiscal\SPED x Espelho NG [Entradas].ahk')
SuspendScript(*) => Suspend(-1)
TabEspelho(*) => Run('Conferir_Espelho_Lançamento\Tabular Espelho Lançamento.ahk')
VerDAPIEntregue(*) => Run('Verificar DAPI Entregue.ahk')
VerRecContrib(*) => Run('Conferir_SPED_Contribuições\VerificarRecEFDContribuicoes.ahk')
VerRecSintegra(*) => Run('Conferir_Arquivos_Sintegra\VerificarRecSintegra.ahk')
VerRecSpedFiscal(*) => Run('Conferir_SPED_Fiscal\VerificarRecSpedFiscal.ahk')
VisuEspelho(*) => Run('Conferir_Espelho_Lançamento\Visualizador Espelho Lançamento.ahk')
VisuSped(*) => Run('Conferir_SPED_Fiscal\Visualizador Sped Fiscal.ahk')
WinCalc(*) => RunAct('calc.exe')
WinSpy(*) => Run('C:\Program Files\AutoHotkey\WindowSpy.ahk')

InlineCalc(*)
{
    if WinExist('Calculadora Inline ahk_class AutoHotkeyGUI')
         if WinGetMinMax()
            WinRestore(), ControlFocus('Edit1')
        else WinMinimize()
    else
        Run('Calculadora Inline.ahk')
}

CriarPastaMesPassado(*)
{
    path := GetExplorerPath(WinExist('A')) '\' mesPassado('MM-MMM')
    (DirExist(path)) || DirCreate(path)
}

RunEverything(*)
{
    RunAct(A_ProgramFiles '\Everything\Everything.exe', , 1)
    ControlFocus('Edit1', 'ahk_exe Everything.exe')
}

VisuXMLNFe(*)
{
    if WinExist('Visualizador XML NFe ahk_class AutoHotkeyGUI')
    {
        WinActivate()
        return
    }
    folder := (A_ThisHotkey = '~Alt & v') ? LastHwnd() : '\\srvrg-saas\rede\PEDRO\XML'
    Run(Format('{} "{}\Conferir_arquivo_XML\Visualizador XML NFe.ahk" "{}"', A_AhkPath, A_WorkingDir, folder))
}

ProtocoloAut(*)
{
    Run(Format('{} "{}\Protocolos\Protocolo Automatico.ahk" "\\srvrg-saas\rede\PEDRO\Impostos {}"', A_AhkPath, A_WorkingDir, mesPassado('MMMM yyyy')))
}

DeletarArquivos(*)
{
    loop files '\\srvrg-saas\rede\PEDRO\Meus Arquivos Magnéticos\*.*'
        if A_LoopFileExt = 'txt' or A_LoopFileExt = 'tmp' or A_LoopFileExt = 'zip'
            try FileDelete(A_LoopFileFullPath)
            catch
                alert('Não foi possível excluir o arquivo:`n' A_LoopFileName
                    . '`nVerifique se o arquivo está em uso', 'Erro T5')

    loop files 'C:\Users\' A_UserName '\Downloads\*.*'
        try FileDelete(A_LoopFileFullPath)
        catch
            alert('Não foi possível excluir o arquivo:`n' A_LoopFileName
                . '`nVerifique se o arquivo está em uso', 'Erro T5')

    alert('Arquivos excluidos com sucesso', 'Ok N T1')
}


CalcularClipBoard(*)
{
    expr := Trim(RegExReplace(A_Clipboard, '(\r|\n)+', '+'), '+')
    expr := RegExReplace(expr, '[^\d\,\+\-\*\/\(\)\s]')
    SendEvent(milhar(Calcular(expr)))
}

AreaRemota(*)
{
    WinSetStyle(-0xC00000, 'ahk_exe mstsc.exe')
    WinMove(40, 1, 1295, 735, 'ahk_exe mstsc.exe')
}

; Fecha todas as guias abertas do explorer de uma vez
FecharExplorer(*)
{
    GroupAdd('ExplorerGroup', 'ahk_class CabinetWClass')
    WinClose('ahk_group ExplorerGroup')
}

/**
 * Seleciona uma pasta com os XML em Downloads para movê-los
 * para a pasta temporária e apaga a pasta em seguida
 */
MoverPastaXML(*)
{
    if not Folder := FileSelect('D', A_MyDocuments '\Hub Sieg', 'Selecione a pasta a ser movida')
        return

    FileDelete('\\srvrg-saas\rede\PEDRO\XML\*.*')

    loop files Folder '\*.*'
        FileMove(A_LoopFileFullPath, '\\srvrg-saas\rede\PEDRO\XML')

    DirDelete(Folder)
    alert('Pasta movida com sucesso', 'Ok N T1')
    VisuXMLNFe()
    WinWait('Visualizador XML NFe ahk_class AutoHotkeyGUI')
    ControlClick('Button3', 'Visualizador XML NFe ahk_class AutoHotkeyGUI',,,, 'NA')
}


; Envia algum arquivo para meu Bot no Telegram
FuncTeleBot(*)
{
    if not SelectFile := FileSelect(3, LastHwnd(), 'Selecione o arquivo para enviar ao BOT')
        return

    if Telebot.SendFile(SelectFile) = 0
        alert('Arquivo enviado com sucesso', 'Ok T1')
    else
        alert('Erro ao enviar arquivo', 'Erro T1')
}

/**
 * Calculadora para inserir resultados em campos de edição
 */
TooltipCalc(*)
{
    ToolTipOptions.SetTitle('Calculadora ToolTip')
    ToolTipOptions.SetColors(0xE4F9E0)
    CalcHook := InputHook('I2', '{Enter}{NumpadEnter}{Esc}')
    UpdateTooltip()
    CalcHook.Start()
    CalcHook.KeyOpt('{Backspace}', 'N')
    CalcHook.OnKeyDown := UpdateTooltip
    CalcHook.OnChar := UpdateTooltip
    CalcHook.Wait()
    CalcHook.Stop()
    ToolTipOptions.SetTitle()
    ToolTipOptions.SetColors(0xFFCC00)
    ToolTip()

    if CalcHook.EndKey = 'Enter' or CalcHook.EndKey = 'NumpadEnter'
        SendEvent(milhar(Calcular(CalcHook.Input)))
    else
        return

    UpdateTooltip(*) => ToolTip('= ' CalcHook.Input)
}

ControlSetEverything(query)
{
    ControlSetText(query, 'Edit1', 'ahk_exe Everything.exe')
    ControlSend('{End}', 'Edit1', 'ahk_exe Everything.exe')
}

AlternarAreaTrabalhoRemota(*)
{
    static Win := WinActive('A')
    if WinActive('A') != NG := WinExist('ahk_exe mstsc.exe')
    {
        Win := WinActive('A')
        WinActivate(NG)
    }
    else
        WinActivate(Win)
}