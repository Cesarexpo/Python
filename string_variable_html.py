a = "https://www.raspberrypi.org/forums/memberlist.php?mode=viewprofile&u=17787&sid=dc679e29cb6b5d2d409b0361c882b579"

message = """<html>
<head></head>
<body><img src="{URL}"></body>
</html>"""

new_message = message.format(URL=a)
# I've put the formatted message in a new variable
# so you can reuse "message" as a template
print(new_message)

code = "We Say Thanks!"
html = """\
<html>
  <head></head>
  <body>
    <p>Thank you for being a loyal customer.<br>
       Here is your unique code to unlock exclusive content:<br>
       <br><br><h1>{code}</h1><br>
       <img src="http://domain.com/footer.jpg">
    </p>
  </body>
</html>
""".format(code=code)
print(html)
