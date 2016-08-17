#include <Base64.h>
#include <ArduinoHttpClient.h>

/*
This example creates a connection to Office365 and Send a Email
via SOAP

Uses code from the following examples

https://www.arduino.cc/en/Tutorial/WiFiWebClient


*/

#include <SPI.h>
#include <WiFi101.h>

char ssid[] = "SSID"; //  your network SSID (name)
char pass[] = "password";    // your network password (use for WPA, or use as key for WEP)
int keyIndex = 0;            // your network key Index number (needed only for WEP)

//Office365 Credentials
String ExUserName = "user@domain.onmicrosoft.com";
String ExPassword = "password";

//Message Details
String Auth = ExUserName + ":" + ExPassword;
String Subject = "Subject of the Message";
String To = "user@domain.com";
String Body = "Something happening in the Body";



int cCount = 0;

int status = WL_IDLE_STATUS;
WiFiSSLClient client;

void setup() {
  //Initialize serial and wait for port to open:
  Serial.begin(9600);
  while (!Serial) {
    ; // wait for serial port to connect. Needed for native USB port only
  }

  // check for the presence of the shield:
  if (WiFi.status() == WL_NO_SHIELD) {
    Serial.println("WiFi shield not present");
    // don't continue:
    while (true);
  }

  // attempt to connect to Wifi network:
  while (status != WL_CONNECTED) {
    Serial.print("Attempting to connect to SSID: ");
    Serial.println(ssid);
    // Connect to WPA/WPA2 network. Change this line if using open or WEP network:
    status = WiFi.begin(ssid, pass);

    // wait 10 seconds for connection:
    delay(10000);
  }
  Serial.println("Connected to wifi");
  printWifiStatus();
  SendMessageEWS(Auth,To,Subject,Body);

}


void SendMessageEWS(String AuthString,String MessageTo, String MessageSubject, String MessageBody)
{
    //Auth Header
    char basicHeader[AuthString.length()+1];
    AuthString.toCharArray(basicHeader,AuthString.length()+1);
    int inputStringLength = sizeof(basicHeader); 
    int encodedLength = Base64.encodedLength(inputStringLength); 
    char encodedString[encodedLength]; 
    Base64.encode(encodedString, basicHeader, inputStringLength); 
    String AuthHeader = "Basic ";

    //EWS SOAP Request
    String content = "<?xml version=\"1.0\" encoding=\"utf-8\"?>";
    content += "<soap:Envelope xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\" xmlns:m=\"http://schemas.microsoft.com/exchange/services/2006/messages\" xmlns:t=\"http://schemas.microsoft.com/exchange/services/2006/types\" xmlns:soap=\"http://schemas.xmlsoap.org/soap/envelope/\">";
    content += "<soap:Header>";
    content += "<t:RequestServerVersion Version=\"Exchange2013_SP1\" />";
    content += "</soap:Header>";
    content += "<soap:Body>";
    content += "<m:CreateItem MessageDisposition=\"SendAndSaveCopy\">";
    content += "<m:SavedItemFolderId>";
    content += "<t:DistinguishedFolderId Id=\"sentitems\" />";
    content += "</m:SavedItemFolderId>";
    content += "<m:Items>";
    content += "<t:Message>";
    content += "<t:Subject>" + MessageSubject + "</t:Subject>";
    content += "<t:Body BodyType=\"HTML\">" + MessageBody + "</t:Body>";
    content += "<t:ToRecipients>";
    content += "<t:Mailbox>";
    content += "<t:EmailAddress>" + MessageTo + "</t:EmailAddress>";
    content += "</t:Mailbox>";
    content += "</t:ToRecipients>";
    content += "</t:Message>";
    content += "</m:Items>";
    content += "</m:CreateItem>";
    content += "</soap:Body>";
    content += "</soap:Envelope>";

    
    Serial.println("\nStarting connection to server...");
    // if you get a connection, report back via serial:
    if (client.connectSSL("outlook.office365.com", 443)) {
      Serial.println("connected to server");
      client.print("POST ");
      client.print("https://outlook.office365.com/EWS/Exchange.asmx");
      client.println(" HTTP/1.1"); 
      client.print("Host: "); 
      client.println("outlook.office365.com");
      client.print("Authorization: Basic ");
      client.println(encodedString); 
      client.println("Connection: close");
      client.print("Content-Type: ");
      client.println("text/xml");
      client.println("User-Agent: mrk1000Sender");
      client.print("Content-Length: ");
      client.println(content.length());
      client.println();
      client.println(content);      
    }
  
}

void loop() {
  // if there are incoming bytes available
  // from the server, read them and print them:
  while (client.available()) {
    char c = client.read();
    Serial.print(c); 
  }

  // if the server's disconnected, stop the client:
  if (!client.connected()) {
    Serial.println();
    Serial.println("disconnecting from server.");
    client.stop();
    // do nothing forevermore:
    while (true);
  }
}


void printWifiStatus() {
  // print the SSID of the network you're attached to:
  Serial.print("SSID: ");
  Serial.println(WiFi.SSID());

  // print your WiFi shield's IP address:
  IPAddress ip = WiFi.localIP();
  Serial.print("IP Address: ");
  Serial.println(ip);

  // print the received signal strength:
  long rssi = WiFi.RSSI();
  Serial.print("signal strength (RSSI):");
  Serial.print(rssi);
  Serial.println(" dBm");
}
