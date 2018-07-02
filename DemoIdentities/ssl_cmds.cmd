set OPENSSL_CONF=C:\OpenSSL-Win32\bin\openssl.cfg

set /p name="name:"
openssl ecparam -genkey -name prime256v1 -out %name%.key
openssl req -new -nodes -key %name%.key -out %name%.csr -subj "/O=Secretarium/CN=%name%/emailAddress=email@fake.com"
openssl req -x509 -nodes -days 3650 -key %name%.key -in %name%.csr -out %name%.crt
openssl pkcs12 -export -out %name%.pfx -inkey %name%.key -in %name%.crt -password pass:%name%
openssl x509 -in %name%.crt -text -noout
openssl verify %name%.crt