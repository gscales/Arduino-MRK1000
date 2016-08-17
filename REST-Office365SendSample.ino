#include <ArduinoJson.h>

#include <Base64.h>
#include <ArduinoHttpClient.h>

/*
This example creates a client object that connects and transfers
data using always SSL.

It is compatible with the methods normally related to plain
connections, like client.connect(host, port).

Written by Arturo Guadalupi
last revision November 2015

*/

#include <SPI.h>
#include <WiFi101.h>

char ssid[] = "SSOecure"; //  your network SSID (name)
char pass[] = "pass@#";    // your network password (use for WPA, or use as key for WEP)
int keyIndex = 0;            // your network key Index number (needed only for WEP)
const size_t MAX_CONTENT_SIZE = 5120;

//Office365 Credentials
String ExUserName = "user@domain.com";
String ExPassword = "passw";
String ClientId = "8fe353d6-efa0-4b0f-aafb-ab7cf3a9b307";
//Message Details
String Auth = ExUserName + ":" + ExPassword;
String Subject = "Subject of the Message";
String To = "mailbox@domain.com";
String Body = "Something happening in the Body";
String Access_Token = "";

bool DebugResponse = false;

int cCount = 0;

int status = WL_IDLE_STATUS;
// if you don't want to use DNS (and reduce your sketch size)
// use the numeric IP instead of the name for the server:

// Initialize the Ethernet client library
// with the IP address and port of the server
// that you want to connect to (port 80 is default for HTTP):
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
  TokenAuth(ClientId,ExUserName,ExPassword);
  if(Access_Token.length() > 0){
      Serial.println("Send Message");
      SendRest(Access_Token,To,Subject,Body);
      Serial.println("Done");
  }

}

void TokenAuth(String ClientId,String UserName, String Password)
{
     char endOfHeaders[] = "\r\n\r\n";
     char passwordCA[Password.length()+1];
     Password.toCharArray(passwordCA,Password.length()+1);
     String content = "resource=https%3A%2F%2Foutlook.office.com&client_id=" + ClientId + "&grant_type=password&username=" + ExUserName + "&password=" + URLEncode(passwordCA) + "&scope=openid";
     Serial.println("\nStarting connection to server...");
     if (client.connectSSL("login.windows.net", 443)) {
       Serial.println("connected to server");
       client.print("POST ");
       client.println("https://login.windows.net/Common/oauth2/token HTTP/1.1");
       client.println("Content-Type: application/x-www-form-urlencoded");
       client.println("client-request-id: " + ClientId);
       client.println("return-client-request-id: true");
       client.println("x-client-CPU: x32");
       client.println("x-client-OS: Arduino");
       client.println("Host: login.windows.net");
       client.print("Content-Length: ");
       client.println(content.length());
       client.println("Expect: 100-continue");
       client.println(""); 
       client.println(content);
       client.find(endOfHeaders);
       bool ok =  client.find(endOfHeaders);
       if (!ok) {
         Serial.println("No response or invalid response!");
       }
       else{
         Serial.println("Request Okay");
       }
       char response[MAX_CONTENT_SIZE];
       readAuthReponse(response, sizeof(response));
    }
}

void readAuthReponse(char* content, size_t maxSize) {
  size_t length = client.readBytes(content, maxSize);
  content[length] = 0;
  Serial.println(content);
  content[length] = 0;
  DynamicJsonBuffer jsonBuffer;
  JsonObject&root = jsonBuffer.parseObject(content);
  if (!root.success()) {
    Serial.println("JSON parsing failed!");
  }
  else{
      String token = root["access_token"]; 
      Access_Token = token;
  }
  
}


String URLEncode(const char* msg)
{
    const char *hex = "0123456789abcdef";
    String encodedMsg = "";

    while (*msg!='\0'){
        if( ('a' <= *msg && *msg <= 'z')
                || ('A' <= *msg && *msg <= 'Z')
                || ('0' <= *msg && *msg <= '9') ) {
            encodedMsg += *msg;
        } else {
            encodedMsg += '%';
            encodedMsg += hex[*msg >> 4];
            encodedMsg += hex[*msg & 15];
        }
        msg++;
    }
    return encodedMsg;
}


void SendRest(String Bearer,String MessageTo, String MessageSubject, String MessageBody)
{
    //DebugResponse = true;
    char endOfHeaders[] = "\r\n\r\n";  
    String content = "{";
    content += "        \"Message\": {";  
    content += "           \"Subject\": \"" + MessageSubject + "\",";  
    content += "            \"Body\": {";  
    content += "                \"ContentType\": \"Text\",";  
    content += "                \"Content\": \"" + MessageBody + "\"";  
    content += "                       },";  
    content += "            \"ToRecipients\": [";  
    content += "                {";  
    content += "                    \"EmailAddress\": {";  
    content += "                        \"Address\": \"" + MessageTo + "\"";  
    content += "                    }";  
    content += "                }";  
    content += "            ]";  
    content += "        },";  
    content += "        \"SaveToSentItems\": \"false\"";  
    content += "    }";  
    Serial.println("\nStarting connection to server...");
    // if you get a connection, report back via serial:
    if (client.connectSSL("outlook.office365.com", 443)) {
      Serial.println("connected to server");
      client.print("POST ");
      client.print("https://outlook.office365.com/api/v2.0/me/sendmail");
      client.println(" HTTP/1.1"); 
      client.print("Host: "); 
      client.println("outlook.office365.com");
      client.print("Authorization: Bearer ");
      client.println(Bearer); 
      client.println("Connection: close");
      client.print("Content-Type: ");
      client.println("application/json");
      client.println("User-Agent: mrk1000Sender");
      client.print("Content-Length: ");
      client.println(content.length());
      client.println();
      client.println(content);
      char okayString[] = "HTTP/1.1 202 Accepted";
      bool ok =  client.find(okayString);
      if (!ok) {
        Serial.println("Message Sent");
      }
      else{
        Serial.println("Request Failed");
      }
    }

}

void loop() {
  // if there are incoming bytes available
  // from the server, read them and print them:
  if(DebugResponse){
  while (client.available()) {
    char c = client.read();
    Serial.print(c); 
  }
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
