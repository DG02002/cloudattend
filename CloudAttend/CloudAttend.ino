#include <Arduino.h>
#include <ESP8266WiFi.h>
#include <ESP8266HTTPClient.h>
#include <WiFiClientSecureBearSSL.h>
#include <SPI.h>
#include <MFRC522.h>
#include <Wire.h>
#include <LiquidCrystal_I2C.h>
#include <time.h>

// ----- CloudAttend configuration -----
constexpr char APPS_SCRIPT_URL[] = "https://script.google.com/macros/s/AKfycbx86ynSP7SkxUqUT98B8Cm2QkeP18NDBPCKWXf5pumFdCCBruYLsTZzswLJ8kZ4vQQ3/exec";

// Wi-Fi hotspots attempted in priority order (shared password).
constexpr char WIFI_PASSWORD[] = "Admin@12345";
constexpr const char *WIFI_SSIDS[] = {
    "Darshans-Phone",
    "Shreys-Phone",
    "Sumits-Phone",
    "Vaibhavs-Phone"};
constexpr size_t WIFI_NETWORK_COUNT = sizeof(WIFI_SSIDS) / sizeof(WIFI_SSIDS[0]);

// ----- Hardware pin assignment (adapt to your wiring) -----
constexpr uint8_t RFID_SS_PIN = D4;
constexpr uint8_t RFID_RST_PIN = D3;
constexpr int BUZZER_PIN = D8; // Set to -1 if no buzzer is attached.

// ----- Wi-Fi and time configuration -----
constexpr unsigned long WIFI_RETRY_DELAY_MS = 3000;
constexpr uint8_t WIFI_MAX_RETRIES = 10;
constexpr long GMT_OFFSET_SEC = 0; // Adjust for your timezone if needed.
constexpr int DAYLIGHT_OFFSET_SEC = 0;
constexpr char NTP_SERVER[] = "pool.ntp.org";

// ----- HTTP configuration -----
constexpr uint8_t MAX_POST_RETRIES = 3;
constexpr uint16_t HTTP_TIMEOUT_MS = 8000;

// ----- LCD configuration -----
LiquidCrystal_I2C lcd(0x27, 16, 2); // Change I2C address if required.
MFRC522 rfid(RFID_SS_PIN, RFID_RST_PIN);

// Forward declarations.
void connectToWiFi();
void ensureTimeSync();
String buildIsoTimestamp();
String buildPayload(const String &uid, const String &isoTimestamp);
String readUidHex(const MFRC522::Uid &uidStruct);
void displayStatus(const String &line1, const String &line2 = "");
void buzz(uint16_t durationMs, uint8_t repeat = 1, uint16_t gapMs = 120);

struct ApiResponse
{
    bool success = false;
    String action;
    int httpCode = 0;
    String message;
};

ApiResponse postScanEvent(const String &uidHex, const String &isoTimestamp);
ApiResponse parseApiResponse(int httpCode, const String &body);

void setup()
{
    Serial.begin(115200);
    delay(200);
    Serial.println();
    Serial.println(F("CloudAttend booting..."));

    if (BUZZER_PIN >= 0)
    {
        pinMode(BUZZER_PIN, OUTPUT);
        digitalWrite(BUZZER_PIN, LOW);
    }

    lcd.init();
    lcd.backlight();
    displayStatus("CloudAttend", "Starting...");

    SPI.begin();
    rfid.PCD_Init();
    delay(50);

    WiFi.mode(WIFI_STA);
    connectToWiFi();

    configTime(GMT_OFFSET_SEC, DAYLIGHT_OFFSET_SEC, NTP_SERVER);
    ensureTimeSync();

    displayStatus("CloudAttend", "Scan your card");
    Serial.println(F("Setup complete."));
}

void loop()
{
    if (WiFi.status() != WL_CONNECTED)
    {
        displayStatus("WiFi lost", "Reconnecting...");
        connectToWiFi();
        displayStatus("CloudAttend", "Scan your card");
    }

    if (!rfid.PICC_IsNewCardPresent() || !rfid.PICC_ReadCardSerial())
    {
        delay(50);
        return;
    }

    const String uidHex = readUidHex(rfid.uid);
    Serial.printf("Card detected: %s\n", uidHex.c_str());
    displayStatus("Card detected", uidHex);
    buzz(60, 1);

    ensureTimeSync();
    const String isoTimestamp = buildIsoTimestamp();
    const ApiResponse response = postScanEvent(uidHex, isoTimestamp);

    if (!response.success)
    {
        Serial.printf("API error (%d): %s\n", response.httpCode, response.message.c_str());
        displayStatus("Send failed", response.message);
        buzz(400, 2);
    }
    else
    {
        if (response.action == "checkin")
        {
            displayStatus("Check-In OK", uidHex);
            buzz(80, 2, 80);
        }
        else if (response.action == "checkout")
        {
            displayStatus("Check-Out OK", uidHex);
            buzz(80, 3, 60);
        }
        else if (response.action == "unregistered")
        {
            displayStatus("Unknown card", uidHex);
            buzz(600, 1);
        }
        else
        {
            displayStatus("Unknown reply", response.action);
            buzz(300, 1);
        }
    }

    rfid.PICC_HaltA();
    rfid.PCD_StopCrypto1();
    delay(1200); // Debounce further reads.
    displayStatus("CloudAttend", "Scan your card");
}

void connectToWiFi()
{
    for (size_t networkIdx = 0; networkIdx < WIFI_NETWORK_COUNT; networkIdx++)
    {
        const char *ssid = WIFI_SSIDS[networkIdx];
        Serial.printf("Attempting SSID: %s\n", ssid);
        displayStatus("WiFi connect", ssid);

        WiFi.begin(ssid, WIFI_PASSWORD);
        uint8_t attempts = 0;
        while (WiFi.status() != WL_CONNECTED && attempts < WIFI_MAX_RETRIES)
        {
            attempts++;
            Serial.printf("  Try %u/%u...\n", attempts, WIFI_MAX_RETRIES);
            delay(WIFI_RETRY_DELAY_MS);
        }

        if (WiFi.status() == WL_CONNECTED)
        {
            Serial.printf("Connected to %s. IP: %s\n", ssid, WiFi.localIP().toString().c_str());
            displayStatus("WiFi OK", WiFi.localIP().toString());
            buzz(70, 2, 50);
            return;
        }

        Serial.printf("Failed to join %s.\n", ssid);
        WiFi.disconnect(true);
        delay(500);
    }

    Serial.println(F("All WiFi attempts failed."));
    displayStatus("WiFi failed", "Check hotspots");
    buzz(500, 1);
}

void ensureTimeSync()
{
    time_t now = time(nullptr);
    uint8_t guard = 0;
    while (now < 1600000000 && guard < 60)
    { // Roughly before 2020-09-13
        delay(500);
        now = time(nullptr);
        guard++;
    }
    if (guard >= 60)
    {
        Serial.println(F("NTP sync timeout."));
    }
}

String buildIsoTimestamp()
{
    time_t now = time(nullptr);
    struct tm timeInfo;
    gmtime_r(&now, &timeInfo);
    char buffer[25];
    strftime(buffer, sizeof(buffer), "%Y-%m-%dT%H:%M:%SZ", &timeInfo);
    return String(buffer);
}

String buildPayload(const String &uid, const String &isoTimestamp)
{
    String payload = "{\"uid\":\"";
    payload += uid;
    payload += "\",\"action\":\"scan\",\"timestamp\":\"";
    payload += isoTimestamp;
    payload += "\"}";
    return payload;
}

String readUidHex(const MFRC522::Uid &uidStruct)
{
    String hex = "";
    for (byte i = 0; i < uidStruct.size; i++)
    {
        if (uidStruct.uidByte[i] < 0x10)
        {
            hex += '0';
        }
        hex += String(uidStruct.uidByte[i], HEX);
    }
    hex.toUpperCase();
    return hex;
}

void displayStatus(const String &line1, const String &line2)
{
    lcd.clear();
    lcd.setCursor(0, 0);
    lcd.print(line1.substring(0, 16));
    lcd.setCursor(0, 1);
    lcd.print(line2.substring(0, 16));
}

void buzz(uint16_t durationMs, uint8_t repeat, uint16_t gapMs)
{
    if (BUZZER_PIN < 0)
    {
        return;
    }
    for (uint8_t i = 0; i < repeat; i++)
    {
        digitalWrite(BUZZER_PIN, HIGH);
        delay(durationMs);
        digitalWrite(BUZZER_PIN, LOW);
        if (i + 1 < repeat)
        {
            delay(gapMs);
        }
    }
}

ApiResponse postScanEvent(const String &uidHex, const String &isoTimestamp)
{
    ApiResponse response;
    const String payload = buildPayload(uidHex, isoTimestamp);

    for (uint8_t attempt = 0; attempt < MAX_POST_RETRIES; attempt++)
    {
        if (WiFi.status() != WL_CONNECTED)
        {
            connectToWiFi();
            if (WiFi.status() != WL_CONNECTED)
            {
                response.message = "WiFi down";
                continue;
            }
        }

        std::unique_ptr<BearSSL::WiFiClientSecure> client(new BearSSL::WiFiClientSecure());
        client->setInsecure(); // For prototype only; trust Apps Script certificate.

        HTTPClient http;
        http.setTimeout(HTTP_TIMEOUT_MS);
        if (!http.begin(*client, APPS_SCRIPT_URL))
        {
            response.message = "Bad URL";
            http.end();
            delay(600);
            continue;
        }

        http.addHeader("Content-Type", "application/json");
        const int httpCode = http.POST(payload);
        const String body = http.getString();
        http.end();

        response = parseApiResponse(httpCode, body);
        if (response.success)
        {
            return response;
        }

        Serial.printf("Attempt %u failed (%d): %s\n", attempt + 1, response.httpCode, response.message.c_str());
        delay(800);
    }

    return response;
}

ApiResponse parseApiResponse(int httpCode, const String &body)
{
    ApiResponse res;
    res.httpCode = httpCode;

    if (httpCode <= 0)
    {
        res.message = HTTPClient::errorToString(httpCode);
        return res;
    }

    if (httpCode != HTTP_CODE_OK)
    {
        res.message = String("HTTP ") + httpCode;
        return res;
    }

    if (body.length() == 0)
    {
        res.message = "Empty body";
        return res;
    }

    if (body.indexOf("\"status\":\"ok\"") != -1)
    {
        res.success = true;
        if (body.indexOf("\"action\":\"checkin\"") != -1)
        {
            res.action = "checkin";
        }
        else if (body.indexOf("\"action\":\"checkout\"") != -1)
        {
            res.action = "checkout";
        }
        else
        {
            res.action = "";
        }
        return res;
    }

    if (body.indexOf("unregistered") != -1)
    {
        res.success = true;
        res.action = "unregistered";
        return res;
    }

    res.message = "Unexpected body";
    return res;
}
