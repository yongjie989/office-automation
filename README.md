<h1>office-automation</h1>

<p>
   Collection many usually Python functions code for Office Automation.
</p>

<h2>Author</h2>
<ul>
<li>
    Yong Jie Huang<br>
    yongjie989@gmail.com<br>
    https://launchpad.net/~yj.huang<br>
    Create time: 2013-05-22
</li>
</ul>

<h2>outlook.py</h2>
<p>
    Send email through Outlook. <br>
    Requirement library: <br>
    <ul>
        <li>pywin32<br>
        http://sourceforge.net/projects/pywin32/files/?source=navbar
        </li>
    </ul>
</p>
<pre>
# Send to mutiple users can input many email ans separate by ; e.g. to = "user1@example.com;user2@example.com;user3.example.com;"
# If would like send mail to CC. e.g. cc = "user4@example.com"
# If would like send mail to BCC. e.g. bcc = "user5@example.com"
# If not use HTML content. html = False
# ftype is for your attach file, e.g. daily_report.ppt here ftype is "ppt"

sendmail_outlook(
    subject = "",
    to = "Yong Jie Huang <yongjie989@gmail.com>;",
    message = "This is email body for html format...",
    attach = "c:\\automail\\file_you_want_toy_send",
    ftype = "ppt",
    html=True
    )
</pre>


