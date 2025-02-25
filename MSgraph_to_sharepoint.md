From MSgraph to Sharepoint to get the UI, you have to impose as MSTeams, you can do this by call when request for token client_id=1fec8e78-bce4-4aaf-ab1b-5451cc387264

#### Step 1: 
- The identifier “1fec8e78-bce4-4aaf-ab1b-5451cc387264” is the fixed client ID assigned by Microsoft to the Teams mobile/desktop application.
- The “resource” parameter in OAuth is often used as an identifier, and the value “00000003-0000-0ff1-ce00-000000000000” usually corresponds to Microsoft Graph API. This is a well-known resource ID.
  
```
POST /Tenant_ID/oauth2/token HTTP/2
Host: login.microsoftonline.com
User-Agent: python-requests/2.32.3
Accept-Encoding: gzip, deflate, br
Accept: */*
Content-Length: 1173
Content-Type: application/x-www-form-urlencoded

client_id=1fec8e78-bce4-4aaf-ab1b-5451cc387264&grant_type=refresh_token&refresh_token=RF_TOKEN&resource=00000003-0000-0ff1-ce00-000000000000
```

### Step 2: 

- Get cookie: Set-Cookie: SPOIDCRL=

```
POST /_api/SP.OAuth.NativeClient/Authenticate HTTP/2
Host: YOUR_COMPANY.sharepoint.com
User-Agent: Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36 Edg/131.0.0.0 Teams/24295.605.3225.8804/49
Accept-Encoding: gzip, deflate, br
Accept: application/json;odata=verbose
Authorization: Bearer ACCESS_TOKEN_FROM_STEP_1
Collectspperfmetrics: SPSQLQueryCount
X-Featureversion: 2
Content-Type: application/json;odata=verbose
Sec-Fetch-Site: same-origin
Sec-Fetch-Mode: cors
Sec-Fetch-Dest: empty
Accept-Language: en-US,en;q=0.9
Priority: u=1, i
Cookie: FeatureOverrides_experiments=[]
Content-Length: 0
```

### Step 3:

```
GET / HTTP/2
Host: YOUR_COMPANY.sharepoint.com
Cookie: SPOIDCRL=
```


