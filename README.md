# Microsoft Graph OAuth Python (msauthpy)

The sample consists of three main files, along with one requirements file. In order to update the config file, please ensure to include the clientId and tenantId. To authenticate the application, the console application utilizes the deviceauth method, which requires user authentication via the Microsoft Devices page at [deviceauth](https://microsoft.com/devicelogin). For detailed instructions on setting up the Microsoft Entra Application, please refer to this [link](https://learn.microsoft.com/en-us/graph/tutorials/python?tabs=aad&tutorial-step=1).

### config.cfg

```
[azure]
clientId = *****************************
tenantId = *****************************
graphUserScopes = User.Read Mail.Read Mail.Send UserActivity.ReadWrite.CreatedByApp
```

### requirements.txt
```
azure-identity
msgraph-sdk
```

### Console application outputs
```
Please choose one of the following options:
0. Exit
1. Display access token
2. List my inbox
3. Send mail
4. List org. users
5. List my activities
```

<img width="718" alt="image" src="https://github.com/user-attachments/assets/744e1d86-19ef-4d53-a499-031e5b88fd40">
