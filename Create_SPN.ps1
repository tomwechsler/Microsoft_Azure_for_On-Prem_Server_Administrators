#Create a SPN

setspn -s http/app1.corp.int filesrv01

setspn -s http/app2.corp.int corp\svcacc

#Now you can select the Delegation tab in the properties of the svcacc account.
