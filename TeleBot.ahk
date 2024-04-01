/************************************************************************
 * @description Class para enviar e receber mensagens do Bot Telegram
 * @author Pedro Henrique C. Xavier
 * @date 2024-03-04
 * @version 2.1-alpha.8
 ***********************************************************************/

/*
Markdown V2 suport
*bold \*text*
_italic \*text_
__underline__
~strikethrough~
||spoiler||
*bold _italic bold ~italic bold strikethrough ||italic bold strikethrough spoiler||~ __underline italic bold___ bold*
[inline URL](http://www.example.com/)
[inline mention of a user](tg://user?id=123456789)
`inline fixed-width code`
```
pre-formatted fixed-width code block
```
```python
pre-formatted fixed-width code block written in the Python programming language
```

HTML suport

<b>bold</b>, <strong>bold</strong>
<i>italic</i>, <em>italic</em>
<u>underline</u>, <ins>underline</ins>
<s>strikethrough</s>, <strike>strikethrough</strike>, <del>strikethrough</del>
<span class="tg-spoiler">spoiler</span>, <tg-spoiler>spoiler</tg-spoiler>
<b>bold <i>italic bold <s>italic bold strikethrough <span class="tg-spoiler">italic bold strikethrough spoiler</span></s> <u>underline italic bold</u></i> bold</b>
<a href="http://www.example.com/">inline URL</a>
<a href="tg://user?id=123456789">inline mention of a user</a>
<code>inline fixed-width code</code>
<pre>pre-formatted fixed-width code block</pre>
<pre><code class="language-python">pre-formatted fixed-width code block written in the Python programming language</code></pre>
*/

if A_LineFile = A_ScriptFullPath
{
    if not SelectFile := FileSelect(3, , 'Selecione o arquivo para enviar ao BOT')
        return

    if Telebot.SendFile(SelectFile) = 0
        MsgBox('Arquivo enviado com sucesso', 'Iconi T1')
    else
        MsgBox('Erro ao enviar arquivo', 'IconX T1')
}

/**
 * Se comunica com meu Bot do Telegram
 */
class Telebot
{
    static token := '1811125778:AAEkBYXzu0N6yLHbHrsN6UBqd28TgFBHpHE'
    static chatID := '76573038'
    static whr := ComObject('WinHttp.WinHttpRequest.5.1')

    /**
     * Envia uma mensagem para o bot do Telegram
     * @param {string} mensagem Mensagem a ser enviada
     * @param {string} parse_mode MarkdownV2 ou HTML
     */
    static SendMessage(str, parse_mode := 'MarkdownV2')
    {
        str := StrReplace(str, '`n', '%0A') ; Substitui quebras de linha pelo código correspondente
        this.whr.Open('GET', 'https://api.telegram.org/bot' this.token '/sendMessage?chat_id=' this.chatID '&text=' str '&parse_mode=' parse_mode, true)
        this.whr.Send()
        this.whr.WaitForResponse()
        return this.whr.ResponseText
    }
    /**
     * Retorna a última mensagem disponível recebida
     * @returns {string} Mensagem
     */
    static GetMessage()
    {
        offset := 0
        updates := '["update_id","message"]'
        this.whr.Open('GET', 'https://api.telegram.org/bot' this.token '/getUpdates?offset=' offset '&allowed_updates=["update_id","message"]', true)
        this.whr.Send()
        this.whr.WaitForResponse()
        message := this.whr.ResponseText
        RegExMatch(message, '\"update_id\"\:(\d+)', &ofs)
        if message != '{"ok":true,"result":[]}'
        {
            this.whr.Open('GET', 'https://api.telegram.org/bot' this.token '/getUpdates?offset=' ofs[1] + 1, true)
            this.whr.Send()
            this.whr.WaitForResponse()
            return message
        }
        else
            return false
    }
    /**
     * Envia um arquivo (Máx 50MB) para o bot do Telegram
     * @param FileName
     * @returns {number}
     */
    static SendFile(FileName)
    {
        curlcommand := '-F document=@"' FileName '" ' 'https://api.telegram.org/bot' this.token '/sendDocument?chat_id=' this.chatID
        return RunWait(A_LineFile '\..\curl.exe ' curlcommand, , 'hide')
    }
}
