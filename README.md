**Mail açma/kapama işlemlerini entegrasyon haline getirmek için;**

- [Microsoft.Graph](https://www.nuget.org/packages/Microsoft.Graph) servisi kullanıldı.
- Öncelikle [Microsoft Azure](https://portal.azure.com "Azure Portal") girişi yapıyoruz,
- Yeni bir [Azure AD application](https://docs.microsoft.com/en-us/azure/active-directory/develop/howto-create-service-principal-portal) oluşturuyoruz bu bilgileri ve tam yetki içeren kullanıcımızın bilgilerini projenin **web.config** dosyasında ilgili alanlara kaydediyoruz

##### Web.config
```xml
  <appSettings>
    <add key="logURL" value="C:\\Users\\aysenur.kocak\\Desktop\\Log.txt"/>
    <add key="username" value="aysenurkocak@sample.com.tr"/>
    <add key="password" value="akw20*!"/>
    <add key="client_id" value="6656a8ce-c75a-4kr6-8d9c-4fd9cm3af56b"/>
    <add key="client_secret" value="SkYegB7F-gk-4R[okLL43Li:1WDvEdIc"/>
  </appSettings>
```

Fonksiyon [kodları](https://github.com/aysenurkocak/office365_user/blob/master/Office365/MailIntegration.svc.cs)



