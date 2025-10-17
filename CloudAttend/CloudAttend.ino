#include <Arduino.h>
#include <ESP8266WiFi.h>
#include <ESP8266HTTPClient.h>
#include <WiFiClientSecureBearSSL.h>
#include <SPI.h>
#include <MFRC522.h>
#include <Wire.h>
#include <LiquidCrystal_I2C.h>
#include <time.h>
#include <math.h>
#include <LCDGraph.h>
#include <memory>
#include <cstring>
#include <vector>
#include <algorithm>

#define LOG_INFO(fmt, ...) Serial.printf("[INFO] " fmt "\n", ##__VA_ARGS__)
#define LOG_WARN(fmt, ...) Serial.printf("[WARN] " fmt "\n", ##__VA_ARGS__)
#define LOG_ERROR(fmt, ...) Serial.printf("[ERROR] " fmt "\n", ##__VA_ARGS__)

// ----- CloudAttend configuration -----
#define K_APPS_SCRIPT_DEPLOYMENT_ID \
    "AKfycbz0F26gZ5EX5VNtZYW2Tr_gyGVgcEMX0LkXSdf4Q64apiLkEBvbplifFICe1TgEHtTo"

constexpr char kAppsScriptUrl[] =
    "https://script.google.com/macros/s/" K_APPS_SCRIPT_DEPLOYMENT_ID "/exec";

// Wi-Fi hotspots attempted in priority order (shared password).
constexpr char kWifiPassword[] = "Admin@12345";
constexpr const char *kWifiSsids[] = {
    "Darshans-Phone",
    "Shreys-Phone",
    "Sumits-Phone",
    "Vaibhavs-Phone",
    "Someones-Phone"};
constexpr size_t kWifiNetworkCount =
    sizeof(kWifiSsids) / sizeof(kWifiSsids[0]);

// ----- Hardware pin assignment (adapt to your wiring) -----
constexpr uint8_t kRfidSsPin = D8;
constexpr uint8_t kRfidRstPin = D3;
constexpr int kBuzzerPin = D0; // Set to -1 if no buzzer is attached.

// ----- Wi-Fi and time configuration -----
constexpr unsigned long kWifiRetryDelayMs = 3000;
constexpr uint8_t kWifiMaxRetries = 10;
constexpr long kGmtOffsetSec = 19800; // UTC+5:30 for India.
constexpr int kDaylightOffsetSec = 0;
constexpr char kNtpPrimary[] = "pool.ntp.org";
constexpr char kNtpBackup[] = "time.google.com";
constexpr char kNtpTertiary[] = "time1.google.com";

// ----- HTTP configuration -----
constexpr uint8_t kMaxPostRetries = 3;
constexpr uint16_t kHttpTimeoutMs = 8000;

// ----- LCD layout configuration -----
constexpr uint8_t kLcdColumns = 16;
constexpr uint8_t kLcdRows = 2;
constexpr uint8_t kGraphWidth = 8;
constexpr uint8_t kGraphHeight = 1;
constexpr uint8_t kGraphFirstRegister = 0;
constexpr uint8_t kGraphRow = 0;
constexpr uint8_t kTimeColumn = 0;
constexpr uint8_t kTimeWidth = 7;
constexpr uint8_t kTimeGapColumns = 1;
constexpr uint8_t kGraphColumn =
    kTimeColumn + kTimeWidth + kTimeGapColumns;
constexpr uint8_t kStatusWidth = kLcdColumns;
constexpr float kGraphPhaseIncrement = 0.25f;
constexpr uint16_t kGraphMaxValue = 220;
constexpr unsigned long kStatusMessageDelayMs = 900;

// ----- LCD configuration -----
LiquidCrystal_I2C lcd(0x27, 16, 2); // Change I2C address if required.
MFRC522 rfid(kRfidSsPin, kRfidRstPin);
LCDGraph<int16_t, LiquidCrystal_I2C> activityGraph(
    kGraphWidth, kGraphHeight, kGraphFirstRegister);

struct StudentRecord
{
    String uid;
    String firstName;
    String lastName;
};

std::vector<StudentRecord> studentRegistry;

// Forward declarations.
bool connectToWiFi();
void ensureTimeSync();
String buildIsoTimestamp();
String buildPayload(const String &uid, const String &isoTimestamp);
String readUidHex(const MFRC522::Uid &uidStruct);
void displayStatus(const String &line1, const String &line2 = "");
void displayBootTitle();
void displayStatusHold(const String &line1, const String &line2 = "",
                       unsigned long holdMs = kStatusMessageDelayMs);
void showConnectionSplash(const String &message);
void showFullscreenStatus(const String &line1,
                          const String &line2 = "");
void restoreOperationalDisplay();
void clearStatusLine();
void writeLcdRow(uint8_t row, const String &value);
void playBootMelody();
void playToneSequence(const uint16_t *notes, const uint16_t *durations,
                      size_t length, uint16_t gapMs = 50);
void playConnectionChime();
void playScanPromptTone();
void playCheckInChime();
void playCheckOutChime();
void playAlreadyOutTone();
void playUnregisteredChime();
void playErrorTone();
void playGenericConfirmationTone();
void updateClockDisplay(bool force = false);
void updateGraphAnimation();
void renderTime(const String &text);
String composeStatusMessage(const String &primary, const String &secondary);
String tailString(const String &value, uint8_t length = 4);
bool performSanityChecks();
bool verifyAppsScriptEndpoint();
String buildHealthCheckUrl();
bool loadStudentRegistry();
bool parseStudentRegistry(const String &body);
const StudentRecord *findStudentByUid(const String &uid);
String buildRegistryUrl();
bool tryConnectToNetwork(const char *ssid);
void upsertStudentCacheRecord(const String &uid, const String &firstName,
                              const String &lastName = "");
String deriveLastNameFromFullName(const String &fullName,
                                  const String &firstName);

struct ApiResponse
{
    bool success = false;
    String action;
    int httpCode = 0;
    String message;
    unsigned long elapsedMs = 0;
    String firstName;
    String fullName;
};

constexpr unsigned long kClockUpdateInterval = 1000;
unsigned long lastClockUpdate = 0;
constexpr unsigned long kGraphUpdateInterval = 200;
unsigned long lastGraphUpdate = 0;
unsigned long bootStartMs = 0;
bool bootSequenceComplete = false;
bool rosterAvailable = false;
bool servicesAvailable = false;
// Tracks whether the time/graph overlay should update the LCD.
bool operationalDisplayActive = false;

ApiResponse postScanEvent(const String &uidHex, const String &isoTimestamp);
ApiResponse parseApiResponse(int httpCode, const String &body);

void setup()
{
    bootStartMs = millis();
    Serial.begin(115200);
    delay(200);
    Serial.println();
    LOG_INFO("CloudAttend boot sequence starting");
    LOG_INFO("Firmware build %s %s", __DATE__, __TIME__);
    LOG_INFO("Initializing peripherals");

    if (kBuzzerPin >= 0)
    {
        pinMode(kBuzzerPin, OUTPUT);
        digitalWrite(kBuzzerPin, LOW);
    }

    lcd.init();
    lcd.backlight();
    lcd.clear();
    activityGraph.begin(&lcd);
    activityGraph.clear();
    displayBootTitle();
    clearStatusLine();
    displayStatusHold("Getting ready");

    LOG_INFO("Starting SPI and RFID reader");
    SPI.begin();
    rfid.PCD_Init();
    LOG_INFO("MFRC522 firmware version 0x%02X",
             rfid.PCD_ReadRegister(MFRC522::VersionReg));
    delay(50);

    WiFi.mode(WIFI_STA);
    connectToWiFi();

    configTime(kGmtOffsetSec, kDaylightOffsetSec, kNtpPrimary, kNtpBackup,
               kNtpTertiary);
    ensureTimeSync();
    displayStatusHold("Updating clock");
    rosterAvailable = loadStudentRegistry();
    displayStatusHold(rosterAvailable ? "Records updated"
                                      : "No saved records");

    servicesAvailable = performSanityChecks();
    const bool systemsHealthy = rosterAvailable && servicesAvailable;
    displayStatusHold(systemsHealthy ? "Done" : "! Service down");

    playBootMelody();

    const unsigned long bootDuration = millis() - bootStartMs;
    LOG_INFO("Boot sequence completed in %lu ms", bootDuration);

    lcd.clear();
    bootSequenceComplete = true;
    lastClockUpdate = 0;
    lastGraphUpdate = 0;
    activityGraph.clear();
    activityGraph.setRegisters();
    operationalDisplayActive = true;
    activityGraph.display(kGraphColumn, kGraphRow);
    updateClockDisplay(true);
    updateGraphAnimation();

    displayStatus(systemsHealthy ? "Ready for scan" : "Limited mode",
                  systemsHealthy ? "" : "Check logs");
    LOG_INFO("Setup complete. System status: %s",
             systemsHealthy ? "normal" : "degraded");
}

void loop()
{
    updateClockDisplay();
    updateGraphAnimation();

    if (WiFi.status() != WL_CONNECTED)
    {
        LOG_WARN("WiFi connection lost");
        showConnectionSplash("Reconnecting");
        const bool rejoined = connectToWiFi();
        if (!rejoined)
        {
            showConnectionSplash("Offline");
            delay(1200);
            return;
        }

        restoreOperationalDisplay();
        const bool healthy = rosterAvailable && servicesAvailable;
        displayStatus(healthy ? "Ready for scan" : "Limited mode",
                      healthy ? "" : "Check logs");
    }

    if (!rfid.PICC_IsNewCardPresent() || !rfid.PICC_ReadCardSerial())
    {
        delay(50);
        return;
    }

    const String uidHex = readUidHex(rfid.uid);
    const StudentRecord *student = findStudentByUid(uidHex);
    const String personLabel = student ? student->firstName : String("");
    LOG_INFO("Card detected %s [%s]", uidHex.c_str(), personLabel.c_str());
    if (!student)
    {
        LOG_WARN("UID %s not present in roster cache", uidHex.c_str());
    }

    // Immediately show scanning status to avoid appearing frozen.
    const String scanDetail = personLabel.length() ? personLabel : uidHex;
    showFullscreenStatus("Scanning", scanDetail);
    playScanPromptTone();

    ensureTimeSync();
    const String isoTimestamp = buildIsoTimestamp();
    LOG_INFO("UTC timestamp %s", isoTimestamp.c_str());
    const ApiResponse response = postScanEvent(uidHex, isoTimestamp);
    LOG_INFO("Scan submission completed in %lu ms", response.elapsedMs);

    String resolvedFirstName = response.firstName.length() ? response.firstName : personLabel;
    String trimmedFullName = response.fullName;
    trimmedFullName.trim();
    if (!resolvedFirstName.length() && trimmedFullName.length())
    {
        const int spacePos = trimmedFullName.indexOf(' ');
        resolvedFirstName = spacePos > 0 ? trimmedFullName.substring(0, spacePos)
                                         : trimmedFullName;
    }
    resolvedFirstName.trim();

    String resolvedLastName = student ? student->lastName : String("");
    if (!resolvedLastName.length() && trimmedFullName.length())
    {
        resolvedLastName = deriveLastNameFromFullName(trimmedFullName, resolvedFirstName);
    }
    resolvedLastName.trim();

    if (response.success && response.action != "unregistered" &&
        resolvedFirstName.length())
    {
        upsertStudentCacheRecord(uidHex, resolvedFirstName, resolvedLastName);
    }

    const String statusDetail = resolvedFirstName.length() ? resolvedFirstName : uidHex;

    if (!response.success)
    {
        LOG_ERROR("API error %d: %s", response.httpCode,
                  response.message.c_str());
        showFullscreenStatus("Error", response.message);
        playErrorTone();
    }
    else
    {
        if (response.action == "checkin")
        {
            LOG_INFO("Action check-in acknowledged for %s", uidHex.c_str());
            showFullscreenStatus("Check-in", statusDetail);
            playCheckInChime();
        }
        else if (response.action == "checkout")
        {
            LOG_INFO("Action check-out acknowledged for %s", uidHex.c_str());
            showFullscreenStatus("Check-out", statusDetail);
            playCheckOutChime();
        }
        else if (response.action == "alreadyCheckedOut")
        {
            LOG_INFO("Duplicate checkout ignored for %s", uidHex.c_str());
            showFullscreenStatus("Already out", statusDetail);
            playAlreadyOutTone();
        }
        else if (response.action == "unregistered")
        {
            LOG_WARN("Unregistered card %s", uidHex.c_str());
            // Show a single concise message for brand new cards
            showFullscreenStatus("Unknown card", uidHex);
            playUnregisteredChime();
        }
        else
        {
            LOG_WARN("Unexpected action token '%s'", response.action.c_str());
            // Keep the UI clean; show a generic confirmation
            showFullscreenStatus("Recorded", statusDetail);
            playGenericConfirmationTone();
        }
    }

    rfid.PICC_HaltA();
    rfid.PCD_StopCrypto1();
    delay(1200); // Debounce further reads.
    restoreOperationalDisplay();
    const bool healthy = rosterAvailable && servicesAvailable;
    displayStatus(healthy ? "Ready for scan" : "Limited mode",
                  healthy ? "" : "Check logs");
    LOG_INFO("Awaiting next scan");
}

bool connectToWiFi()
{
    displayStatus("Finding networks");
    WiFi.disconnect(true);
    delay(80);

    struct NetworkCandidate
    {
        const char *ssid;
        int32_t rssi;
    };

    std::vector<const char *> priority;
    priority.reserve(kWifiNetworkCount);
    std::vector<NetworkCandidate> detected;

    const int16_t networkCount = WiFi.scanNetworks(false, false);
    if (networkCount > 0)
    {
        detected.reserve(static_cast<size_t>(networkCount));
        for (int16_t idx = 0; idx < networkCount; idx++)
        {
            const String scannedSsid = WiFi.SSID(idx);
            for (size_t knownIdx = 0; knownIdx < kWifiNetworkCount; knownIdx++)
            {
                if (scannedSsid == kWifiSsids[knownIdx])
                {
                    detected.push_back({kWifiSsids[knownIdx], WiFi.RSSI(idx)});
                    break;
                }
            }
        }
    }
    else
    {
        LOG_WARN("WiFi scan returned %d", networkCount);
    }

    WiFi.scanDelete();

    if (!detected.empty())
    {
        std::sort(
            detected.begin(), detected.end(),
            [](const NetworkCandidate &lhs, const NetworkCandidate &rhs)
            {
                return lhs.rssi > rhs.rssi;
            });

        for (const auto &candidate : detected)
        {
            priority.push_back(candidate.ssid);
        }
    }

    for (size_t idx = 0; idx < kWifiNetworkCount; idx++)
    {
        const char *ssid = kWifiSsids[idx];
        if (std::find(priority.begin(), priority.end(), ssid) == priority.end())
        {
            priority.push_back(ssid);
        }
    }

    bool connected = false;
    for (size_t idx = 0; idx < priority.size(); idx++)
    {
        const char *ssid = priority[idx];
        LOG_INFO("Attempting WiFi SSID '%s'", ssid);
        if (tryConnectToNetwork(ssid))
        {
            WiFi.setAutoReconnect(true);
            connected = true;
            break;
        }
    }

    if (!connected)
    {
        LOG_ERROR("All WiFi attempts exhausted without success");
        displayStatusHold("Network unavailable");
        playErrorTone();
    }

    return connected;
}

bool tryConnectToNetwork(const char *ssid)
{
    displayStatusHold("Connecting to");
    displayStatusHold(ssid);
    WiFi.begin(ssid, kWifiPassword);
    uint8_t attempts = 0;

    while (WiFi.status() != WL_CONNECTED && attempts < kWifiMaxRetries)
    {
        attempts++;
        LOG_INFO("WiFi attempt %u/%u RSSI %d", attempts, kWifiMaxRetries,
                 WiFi.RSSI());
        displayStatus("Attempt ",
                      String(attempts) + "/" + String(kWifiMaxRetries));
        updateClockDisplay();
        updateGraphAnimation();
        delay(kWifiRetryDelayMs);
    }

    if (WiFi.status() == WL_CONNECTED)
    {
        const uint8_t attemptCount = attempts ? attempts : 1;
        LOG_INFO("Connected to WiFi SSID '%s' after %u attempt(s)", ssid,
                 attemptCount);
        displayStatusHold("Connected");
        playConnectionChime();
        return true;
    }

    LOG_WARN("Failed to join SSID '%s' (status %d)", ssid, WiFi.status());
    WiFi.disconnect(true);
    displayStatusHold("Still searching");
    delay(400);
    return false;
}

void ensureTimeSync()
{
    time_t now = time(nullptr);
    uint8_t guard = 0;
    LOG_INFO("Waiting for NTP synchronisation");
    while (now < 1600000000 && guard < 60)
    { // Roughly before 2020-09-13
        delay(500);
        now = time(nullptr);
        guard++;
        updateGraphAnimation();
    }
    if (guard >= 60)
    {
        LOG_WARN("NTP synchronisation timed out");
    }
    else
    {
        LOG_INFO("NTP synchronised after %u cycles (epoch %ld)", guard,
                 static_cast<long>(now));
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
    LOG_INFO("Payload %s", payload.c_str());
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

void updateClockDisplay(bool force)
{
    if (!bootSequenceComplete)
    {
        return;
    }
    if (!operationalDisplayActive)
    {
        return;
    }
    const unsigned long nowMs = millis();
    if (!force && (nowMs - lastClockUpdate) < kClockUpdateInterval)
    {
        return;
    }

    time_t rawTime = time(nullptr);
    if (rawTime < 100000)
    {
        return; // NTP not ready yet.
    }

    struct tm timeInfo;
    localtime_r(&rawTime, &timeInfo);

    char buffer[9];
    strftime(buffer, sizeof(buffer), "%I:%M%p", &timeInfo);
    renderTime(String(buffer));
    lastClockUpdate = nowMs;
}

void updateGraphAnimation()
{
    const unsigned long nowMs = millis();
    if (!bootSequenceComplete)
    {
        lastGraphUpdate = nowMs;
        return;
    }
    if (!operationalDisplayActive)
    {
        lastGraphUpdate = nowMs;
        return;
    }
    if ((nowMs - lastGraphUpdate) < kGraphUpdateInterval)
    {
        return;
    }

    static float phase = 0.0f;

    for (uint8_t samples = 0; samples < 2; samples++)
    {
        const float normalized = (sinf(phase) * 0.45f) + 0.5f; // 0.05 .. 0.95 roughly
        const int16_t sample = static_cast<int16_t>(normalized * kGraphMaxValue);
        activityGraph.add(sample);

        phase += kGraphPhaseIncrement;
        if (phase >= TWO_PI)
        {
            phase -= TWO_PI;
        }
    }

    activityGraph.setRegisters();
    activityGraph.display(kGraphColumn, kGraphRow);

    lastGraphUpdate = nowMs;
}

void renderTime(const String &text)
{
    String content = text;
    content.trim();
    if (content.length() > kTimeWidth)
    {
        content = content.substring(0, kTimeWidth);
    }

    lcd.setCursor(kTimeColumn, kGraphRow);
    for (uint8_t i = 0; i < kTimeWidth; i++)
    {
        lcd.print(' ');
    }

    lcd.setCursor(kTimeColumn, kGraphRow);
    lcd.print(content);
    lcd.setCursor(kTimeColumn + kTimeWidth, kGraphRow);
    lcd.print(' ');
}

String tailString(const String &value, uint8_t length)
{
    if (value.length() <= length)
    {
        return value;
    }
    return value.substring(value.length() - length);
}

String composeStatusMessage(const String &primary, const String &secondary)
{
    String message = primary;
    message.trim();

    if (message.length() > kStatusWidth)
    {
        message = message.substring(0, kStatusWidth);
    }

    if (secondary.length() && message.length() < kStatusWidth)
    {
        String detail = secondary;
        detail.trim();
        if (message.length())
        {
            message += ' ';
        }
        const int remaining = kStatusWidth - message.length();
        if (remaining > 0)
        {
            if (detail.length() > remaining)
            {
                detail = detail.substring(0, remaining);
            }
            message += detail;
        }
    }

    while (message.length() < kStatusWidth)
    {
        message += ' ';
    }

    return message;
}

void displayBootTitle()
{
    const String title = "CloudAttend v0.2";
    const uint8_t padding = (title.length() < kLcdColumns)
                                ? (kLcdColumns - title.length()) / 2
                                : 0;
    lcd.setCursor(0, 0);
    for (uint8_t i = 0; i < kLcdColumns; i++)
    {
        lcd.print(' ');
    }
    lcd.setCursor(padding, 0);
    lcd.print(title);
}

void clearStatusLine()
{
    lcd.setCursor(0, 1);
    for (uint8_t i = 0; i < kLcdColumns; i++)
    {
        lcd.print(' ');
    }
    lcd.setCursor(0, 1);
}

void displayStatusHold(const String &line1, const String &line2, unsigned long holdMs)
{
    displayStatus(line1, line2);
    const unsigned long delayMs = holdMs ? holdMs : kStatusMessageDelayMs;
    delay(delayMs);
    if (bootSequenceComplete)
    {
        updateClockDisplay(true);
        updateGraphAnimation();
    }
}

void displayStatus(const String &line1, const String &line2)
{
    const String message = composeStatusMessage(line1, line2);
    lcd.setCursor(0, 1);
    lcd.print(message);
}

void writeLcdRow(uint8_t row, const String &value)
{
    String content = value;
    content.trim();
    if (content.length() > kLcdColumns)
    {
        content = content.substring(0, kLcdColumns);
    }

    lcd.setCursor(0, row);
    lcd.print(content);
    const uint8_t padding = (content.length() < kLcdColumns)
                                ? (kLcdColumns - content.length())
                                : 0;
    for (uint8_t idx = 0; idx < padding; idx++)
    {
        lcd.print(' ');
    }
}

void showFullscreenStatus(const String &line1, const String &line2)
{
    operationalDisplayActive = false;
    lcd.clear();
    writeLcdRow(0, line1);
    writeLcdRow(1, line2);
}

void showConnectionSplash(const String &message)
{
    operationalDisplayActive = false;
    lcd.clear();
    displayBootTitle();
    displayStatus(message);
}

void restoreOperationalDisplay()
{
    lcd.clear();
    operationalDisplayActive = true;
    activityGraph.setRegisters();
    activityGraph.display(kGraphColumn, kGraphRow);
    updateClockDisplay(true);
    updateGraphAnimation();
}

void playToneSequence(const uint16_t *notes, const uint16_t *durations,
                      size_t length, uint16_t gapMs)
{
    if (kBuzzerPin < 0 || !notes || !durations || length == 0)
    {
        return;
    }

    for (size_t idx = 0; idx < length; idx++)
    {
        const uint16_t frequency = notes[idx];
        const uint16_t noteDuration = durations[idx];

        if (frequency > 0)
        {
            tone(kBuzzerPin, frequency, noteDuration);
            delay(noteDuration);
            noTone(kBuzzerPin);
        }
        else
        {
            noTone(kBuzzerPin);
            delay(noteDuration);
        }

        if (gapMs && idx + 1 < length)
        {
            delay(gapMs);
        }
    }
    noTone(kBuzzerPin);
}

void playBootMelody()
{
    const uint16_t notes[] = {523, 659, 784, 1047}; // C5, E5, G5, C6
    const uint16_t durations[] = {220, 180, 220, 360};
    playToneSequence(notes, durations, sizeof(notes) / sizeof(notes[0]), 60);
}

void playConnectionChime()
{
    const uint16_t notes[] = {1047, 1319}; // C6, E6
    const uint16_t durations[] = {160, 200};
    playToneSequence(notes, durations, sizeof(notes) / sizeof(notes[0]), 50);
}

void playScanPromptTone()
{
    const uint16_t notes[] = {1568}; // G6
    const uint16_t durations[] = {100};
    playToneSequence(notes, durations, sizeof(notes) / sizeof(notes[0]), 0);
}

void playCheckInChime()
{
    const uint16_t notes[] = {1319, 1661, 2093}; // E6, G#6, C7
    const uint16_t durations[] = {140, 160, 240};
    playToneSequence(notes, durations, sizeof(notes) / sizeof(notes[0]), 45);
}

void playCheckOutChime()
{
    const uint16_t notes[] = {1760, 1480, 1175}; // A6, F#6, D6
    const uint16_t durations[] = {140, 140, 220};
    playToneSequence(notes, durations, sizeof(notes) / sizeof(notes[0]), 45);
}

void playAlreadyOutTone()
{
    const uint16_t notes[] = {988, 831}; // B5, G#5
    const uint16_t durations[] = {120, 160};
    playToneSequence(notes, durations, sizeof(notes) / sizeof(notes[0]), 40);
}

void playUnregisteredChime()
{
    const uint16_t notes[] = {784, 622, 523}; // G5, D#5, C5
    const uint16_t durations[] = {160, 160, 280};
    playToneSequence(notes, durations, sizeof(notes) / sizeof(notes[0]), 50);
}

void playErrorTone()
{
    const uint16_t notes[] = {392, 330, 262}; // G4, E4, C4
    const uint16_t durations[] = {200, 200, 320};
    playToneSequence(notes, durations, sizeof(notes) / sizeof(notes[0]), 70);
}

void playGenericConfirmationTone()
{
    const uint16_t notes[] = {1175, 1568}; // D6, G6
    const uint16_t durations[] = {140, 200};
    playToneSequence(notes, durations, sizeof(notes) / sizeof(notes[0]), 40);
}

bool loadStudentRegistry()
{
    if (WiFi.status() != WL_CONNECTED)
    {
        LOG_WARN("Roster sync skipped: WiFi offline");
        return false;
    }

    std::unique_ptr<BearSSL::WiFiClientSecure> client(new BearSSL::WiFiClientSecure());
    client->setInsecure();

    HTTPClient http;
    http.setTimeout(kHttpTimeoutMs);
    http.setFollowRedirects(HTTPC_STRICT_FOLLOW_REDIRECTS);
    const String url = buildRegistryUrl();
    LOG_INFO("Requesting roster from %s", url.c_str());
    if (!http.begin(*client, url))
    {
        LOG_ERROR("Roster request initialisation failed");
        http.end();
        return false;
    }

    const int httpCode = http.GET();
    const String body = http.getString();
    http.end();

    if (httpCode < HTTP_CODE_OK || httpCode >= 300)
    {
        LOG_WARN("Roster HTTP %d", httpCode);
        return false;
    }

    const bool parsed = parseStudentRegistry(body);
    if (!parsed)
    {
        LOG_WARN("Roster parse yielded no records");
    }
    else
    {
        LOG_INFO("Roster loaded with %u records", static_cast<unsigned>(studentRegistry.size()));
    }
    return parsed;
}

bool parseStudentRegistry(const String &body)
{
    studentRegistry.clear();
    if (body.length() == 0)
    {
        LOG_WARN("Roster payload empty");
        return false;
    }

    uint16_t imported = 0;
    int start = 0;
    while (start < body.length())
    {
        int end = body.indexOf('\n', start);
        if (end == -1)
        {
            end = body.length();
        }
        String line = body.substring(start, end);
        line.trim();
        start = end + 1;

        if (!line.length() || line.startsWith("#"))
        {
            continue;
        }

        if (line.startsWith("uid") || line.startsWith("UID") || line.startsWith("CARD_UID"))
        {
            continue; // Skip header rows.
        }

        const int firstComma = line.indexOf(',');
        if (firstComma == -1)
        {
            LOG_WARN("Skipping malformed roster row: %s", line.c_str());
            continue;
        }
        const int secondComma = line.indexOf(',', firstComma + 1);

        String uid = line.substring(0, firstComma);
        uid.trim();
        uid.toUpperCase();

        String firstName;
        String lastName;

        if (secondComma == -1)
        {
            firstName = line.substring(firstComma + 1);
            lastName = "";
        }
        else
        {
            firstName = line.substring(firstComma + 1, secondComma);
            lastName = line.substring(secondComma + 1);
        }

        firstName.trim();
        lastName.trim();

        if (!uid.length() || !firstName.length())
        {
            LOG_WARN("Skipping roster row missing uid/name: %s", line.c_str());
            continue;
        }

        studentRegistry.push_back(StudentRecord{uid, firstName, lastName});
        imported++;
    }

    return imported > 0;
}

const StudentRecord *findStudentByUid(const String &uid)
{
    for (size_t idx = 0; idx < studentRegistry.size(); idx++)
    {
        if (studentRegistry[idx].uid == uid)
        {
            return &studentRegistry[idx];
        }
    }
    return nullptr;
}

void upsertStudentCacheRecord(const String &uid, const String &firstName,
                              const String &lastName)
{
    if (!uid.length() || !firstName.length())
    {
        return;
    }

    for (size_t idx = 0; idx < studentRegistry.size(); idx++)
    {
        if (studentRegistry[idx].uid == uid)
        {
            if (studentRegistry[idx].firstName != firstName)
            {
                studentRegistry[idx].firstName = firstName;
            }
            if (lastName.length())
            {
                studentRegistry[idx].lastName = lastName;
            }
            return;
        }
    }

    studentRegistry.push_back(StudentRecord{uid, firstName, lastName});
}

String deriveLastNameFromFullName(const String &fullName,
                                  const String &firstName)
{
    String trimmedFull = fullName;
    trimmedFull.trim();
    if (!trimmedFull.length())
    {
        return String("");
    }

    String trimmedFirst = firstName;
    trimmedFirst.trim();
    if (!trimmedFirst.length())
    {
        const int firstSpace = trimmedFull.indexOf(' ');
        if (firstSpace > 0 && firstSpace + 1 < trimmedFull.length())
        {
            String tail = trimmedFull.substring(firstSpace + 1);
            tail.trim();
            return tail;
        }
        return String("");
    }

    if (trimmedFull.length() <= trimmedFirst.length())
    {
        return String("");
    }

    if (trimmedFull.startsWith(trimmedFirst))
    {
        String tail = trimmedFull.substring(trimmedFirst.length());
        tail.trim();
        return tail;
    }

    return String("");
}

String buildRegistryUrl()
{
    String url(kAppsScriptUrl);
    url += (url.indexOf('?') == -1) ? "?registry=1" : "&registry=1";
    return url;
}

bool performSanityChecks()
{
    displayStatusHold("Starting");
    displayStatusHold("PowerOn SelfTest");
    const bool urlConfigured = strlen(kAppsScriptUrl) > 0;
    // displayStatusHold(urlConfigured ? "Service URL set" : "Service URL missing");
    const unsigned long startMs = millis();
    const bool appsScriptOk = verifyAppsScriptEndpoint();
    displayStatusHold(appsScriptOk ? "Database online" : "Database Down");
    const unsigned long duration = millis() - startMs;
    LOG_INFO("Sanity checks completed in %lu ms", duration);
    return urlConfigured && appsScriptOk;
}

bool verifyAppsScriptEndpoint()
{
    if (WiFi.status() != WL_CONNECTED)
    {
        LOG_WARN("Skipping Apps Script health check (WiFi offline)");
        return false;
    }

    std::unique_ptr<BearSSL::WiFiClientSecure> client(new BearSSL::WiFiClientSecure());
    client->setInsecure();

    HTTPClient http;
    http.setTimeout(kHttpTimeoutMs);
    http.setFollowRedirects(HTTPC_STRICT_FOLLOW_REDIRECTS);
    const String url = buildHealthCheckUrl();
    LOG_INFO("Performing Apps Script health check");
    if (!http.begin(*client, url))
    {
        LOG_ERROR("Health check initialisation failed");
        http.end();
        return false;
    }

    const unsigned long start = millis();
    const int httpCode = http.GET();
    const unsigned long elapsed = millis() - start;
    LOG_INFO("Health check HTTP %d (%lu ms)", httpCode, elapsed);
    http.end();

    return httpCode >= 200 && httpCode < 400;
}

String buildHealthCheckUrl()
{
    String url(kAppsScriptUrl);
    url += (url.indexOf('?') == -1) ? "?health=1" : "&health=1";
    return url;
}

ApiResponse postScanEvent(const String &uidHex, const String &isoTimestamp)
{
    ApiResponse response;
    const String payload = buildPayload(uidHex, isoTimestamp);
    const unsigned long overallStart = millis();

    for (uint8_t attempt = 0; attempt < kMaxPostRetries; attempt++)
    {
        if (WiFi.status() != WL_CONNECTED)
        {
            if (!connectToWiFi())
            {
                response.message = "WiFi down";
                response.httpCode = 0;
                LOG_WARN("Cannot send payload: WiFi disconnected");
                continue;
            }
        }

        std::unique_ptr<BearSSL::WiFiClientSecure> client(new BearSSL::WiFiClientSecure());
        client->setInsecure(); // For prototype only; trust Apps Script certificate.
        LOG_INFO("Starting HTTPS POST attempt %u/%u", attempt + 1,
                 kMaxPostRetries);

        HTTPClient http;
        http.setTimeout(kHttpTimeoutMs);
        http.setFollowRedirects(HTTPC_STRICT_FOLLOW_REDIRECTS);
        if (!http.begin(*client, kAppsScriptUrl))
        {
            response.message = "Bad URL";
            response.httpCode = 0;
            LOG_ERROR("http.begin failed for Apps Script URL");
            http.end();
            delay(600);
            continue;
        }

        http.addHeader("Content-Type", "application/json");
        const unsigned long requestStart = millis();
        const int httpCode = http.POST(payload);
        const unsigned long requestElapsed = millis() - requestStart;
        LOG_INFO("HTTP POST completed with code %d (%lu ms)", httpCode, requestElapsed);
        const String body = http.getString();
        LOG_INFO("Response body length %u", body.length());
        http.end();

        response = parseApiResponse(httpCode, body);
        response.elapsedMs = millis() - overallStart;
        if (response.success)
        {
            LOG_INFO("Server acknowledged action '%s'", response.action.c_str());
            return response;
        }

        LOG_WARN("Attempt %u failed (%d): %s", attempt + 1, response.httpCode, response.message.c_str());
        delay(800);
    }

    response.elapsedMs = millis() - overallStart;
    return response;
}

ApiResponse parseApiResponse(int httpCode, const String &body)
{
    ApiResponse res;
    res.httpCode = httpCode;

    if (httpCode <= 0)
    {
        res.message = HTTPClient::errorToString(httpCode);
        LOG_ERROR("HTTP transport error: %s", res.message.c_str());
        return res;
    }

    if (httpCode != HTTP_CODE_OK)
    {
        res.message = String("HTTP ") + httpCode;
        LOG_WARN("Non-200 response body: %s", body.c_str());
        return res;
    }

    if (body.length() == 0)
    {
        res.message = "Empty body";
        LOG_WARN("Received empty HTTP body");
        return res;
    }

    auto contains = [](const String &haystack, const char *needle)
    {
        return haystack.indexOf(needle) != -1;
    };

    if (contains(body, "\"status\":\"ok\""))
    {
        res.success = true;

        if (contains(body, "\"action\":\"checkin\""))
        {
            res.action = "checkin";
        }
        else if (contains(body, "\"action\":\"checkout\""))
        {
            res.action = "checkout";
        }
        else if (contains(body, "\"action\":\"alreadyCheckedOut\""))
        {
            res.action = "alreadyCheckedOut";
        }
        else if (contains(body, "\"action\":\"unregistered\""))
        {
            res.action = "unregistered";
        }
        else
        {
            res.action = "acknowledged";
            LOG_WARN("Action token missing from ok response: %s", body.c_str());
        }

        // Capture first name for downstream UI hints when available.
        const int keyPos = body.indexOf("\"firstName\":");
        if (keyPos != -1)
        {
            const int quoteStart = body.indexOf('"', keyPos + 12);
            const int quoteEnd = quoteStart != -1 ? body.indexOf('"', quoteStart + 1) : -1;
            if (quoteStart != -1 && quoteEnd != -1 && quoteEnd > quoteStart)
            {
                res.firstName = body.substring(quoteStart + 1, quoteEnd);
            }
        }

        const int fullKeyPos = body.indexOf("\"fullName\":");
        if (fullKeyPos != -1)
        {
            const int quoteStart = body.indexOf('"', fullKeyPos + 11);
            const int quoteEnd = quoteStart != -1 ? body.indexOf('"', quoteStart + 1) : -1;
            if (quoteStart != -1 && quoteEnd != -1 && quoteEnd > quoteStart)
            {
                res.fullName = body.substring(quoteStart + 1, quoteEnd);
            }
        }

        return res;
    }

    if (contains(body, "\"status\":\"error\""))
    {
        res.message = "Script error";
        const int keyPos = body.indexOf("\"message\":");
        if (keyPos != -1)
        {
            const int quoteStart = body.indexOf('"', keyPos + 10);
            const int quoteEnd = quoteStart != -1 ? body.indexOf('"', quoteStart + 1) : -1;
            if (quoteStart != -1 && quoteEnd != -1 && quoteEnd > quoteStart)
            {
                res.message = body.substring(quoteStart + 1, quoteEnd);
            }
        }
        LOG_ERROR("Apps Script reported error: %s", body.c_str());
        return res;
    }

    if (contains(body, "unregistered"))
    {
        res.success = true;
        res.action = "unregistered";
        return res;
    }

    res.message = "Unexpected body";
    LOG_WARN("Unexpected response payload: %s", body.c_str());
    return res;
}
