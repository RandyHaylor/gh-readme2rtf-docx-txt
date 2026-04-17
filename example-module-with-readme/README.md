# HarmonEyes Android SDK

The HarmonEyes Android SDK provides real-time cognitive load scoring and sleepiness detection from eye-tracking data. It supports both generic gaze input (from any eye tracker) and native integration with the Pupil Labs Neon eye tracker.

## Requirements

- Android API 21+ (Lollipop)
- ARM64 device or emulator (`arm64-v8a`)
- Valid HarmonEyes license key
- Eye tracking data to process — either from a live eye tracker or pre-recorded. The SDK accepts gaze data as a real-time stream or in async batches. You will need to translate your eye data into the `GazeSample` format the SDK expects.

## Installation

1. Download the latest AAR from [Releases](https://github.com/RightEyeLLC/harmoneyes-android-sdk/releases)
2. Copy the AAR into your project's `app/libs/` folder
3. Add the dependency to your `app/build.gradle.kts`:

```kotlin
dependencies {
    implementation(files("libs/harmoneyes-android-sdk-1.0.0.aar"))
}
```

4. The SDK requires these permissions (declared in its manifest, merged automatically):

```xml
<uses-permission android:name="android.permission.INTERNET" />
```

- `INTERNET` — for license validation and local caching

## Quick Start

### 1. Configure the session

Create a session config with your license key, model path, and callbacks. Callbacks are invoked on a background thread as results arrive — keep them fast.

```kotlin
val config = SessionConfig(
    licenseKey = "your-license-key",
    cognitiveLoadModelPath = "/path/to/cog_load.json",
    outputDir = context.filesDir.absolutePath,
    cognitiveLoadListener = { prediction ->
        handleCognitiveLoad(prediction)
    },
    sleepinessListener = { prediction ->
        handleSleepiness(prediction)
    },
    sdkLogListener = { message ->
        Log.d("HarmonEyes", message)
    }
)
```

### 2. Validate License and Start a Session

Validate the license key before starting a session for real work:

```kotlin
// Create a temporary handle to validate
val tempHandle = HarmonEyesSDK.create(config)
val licenseInfo = HarmonEyesSDK.licenseInfo(tempHandle)
HarmonEyesSDK.destroy(tempHandle)

// Check if valid
if (licenseInfo?.status?.uppercase()?.contains("VALID") != true) {
    Log.e("HarmonEyes", "Invalid license: ${licenseInfo?.status}")
    return
}

// License is valid — create a new handle for the session
val handle = HarmonEyesSDK.create(config)
HarmonEyesSDK.startSession(handle, "session-001", "reading")
```

### 3. Feed gaze samples

Feed samples continuously as they arrive from your eye tracker. Results are generated automatically once enough samples accumulate.

```kotlin
val samples = listOf(
    GazeSample(
        timestampNanoseconds = System.nanoTime(),
        leftEyeX = 800f, leftEyeY = 600f,
        rightEyeX = 810f, rightEyeY = 605f,
        pupilDiameterLeft = 4.2, pupilDiameterRight = 4.1,
        pupilDiameterCombined = 4.15,
        blinkId = -1
    )
)
// Submits to a queue for processing; callbacks will trigger when processing completes
HarmonEyesSDK.addGazeSamples(handle, samples)
```

### 4. Poll results directly (optional)

You can also query the latest result at any time without waiting for callbacks:

```kotlin
val cogLoad = HarmonEyesSDK.getCogLoad(handle)
val sleepiness = HarmonEyesSDK.getSleepiness(handle)
```

Using the callbacks (from step 1) to record results is the better practice — it ensures you capture every result as soon as it's ready, rather than missing updates or polling at inopportune times.

### 5. Format results for display

When you need to show results in the UI, format the raw data:

```kotlin
private fun formatCognitiveLoad(prediction: CognitiveLoadPrediction): String {
    return buildString {
        appendLine("Load: ${prediction.level}")  // Low, Moderate, High
        appendLine("Confidence: ${(prediction.confidence * 100).toInt()}%")
        appendLine("High Load Probability: ${(prediction.probabilityOfHighLoad * 100).toInt()}%")
        appendLine("Batch #${prediction.batchNumber}")
    }
}

private fun formatSleepiness(prediction: SleepinessResult): String {
    return buildString {
        appendLine("Alertness: ${prediction.level}")  // Alert, NeitherAlertNorDrowsy, RatherDrowsy, Drowsy
        appendLine("Confidence: ${(prediction.confidence * 100).toInt()}%")
        appendLine("Batch #${prediction.batchNumber}")
    }
}

// Use in callbacks or UI code
if (lastCogLoad != null) {
    statusText.text = formatCognitiveLoad(lastCogLoad!!)
}
```

### 6. Store or Act on Results

Save incoming predictions and use them to drive decision logic and app behavior:

```kotlin
private var lastCogLoad: CognitiveLoadPrediction? = null
private var lastSleepiness: SleepinessResult? = null

private fun handleCognitiveLoad(prediction: CognitiveLoadPrediction) {
    lastCogLoad = prediction  // Store for later use
    File(context.filesDir, "cogLoad.json").appendText("$prediction\n")
    
    if (prediction.confidence > 0.85) {
        // High confidence — safe to act on the prediction
    }
    if (prediction.level == CognitiveLoadLevel.High) {
        // Suggest a break, reduce visual complexity, etc.
    }
}

private fun handleSleepiness(prediction: SleepinessResult) {
    lastSleepiness = prediction  // Store for later use
    File(context.filesDir, "sleepiness.json").appendText("$prediction\n")
    
    if (prediction.level == SleepinessLevel.Drowsy && prediction.confidence > 0.8) {
        // Alert user, disable safety-critical operations, etc.
    }
}
```

### 7. Stop and clean up

Always call `stopSession` before `destroy`. `stopSession` drains any in-flight samples and waits for the native engine to finish processing before releasing resources.

```kotlin
HarmonEyesSDK.stopSession(handle)
HarmonEyesSDK.destroy(handle)
```

## API Reference

All SDK functions are on the `HarmonEyesSDK` object. Sessions and Neon clients are identified by a `Long` handle.

### Session

`create(config)` — Create a session, returns handle
- `config: SessionConfig` — license key, model path, output dir, display config, listeners
- returns `Long` — session handle e.g. `140532891648`
```kotlin
val handle = HarmonEyesSDK.create(config)
```

`startSession(handle, sessionId, taskType)` — Begin a named session
- `handle: Long` — session handle from `create()` e.g. `140532891648`
- `sessionId: String` — unique session identifier e.g. `"session-001"`
- `taskType: String` — task label e.g. `"reading"`
- returns `Boolean` — true on success e.g. `true`
```kotlin
HarmonEyesSDK.startSession(handle, "session-001", "reading")
```

`addGazeSamples(handle, samples)` — Feed gaze data into the session
- `handle: Long` — session handle e.g. `140532891648`
- `samples: List<GazeSample>` — gaze data points e.g. `listOf(gazeSample)`
```kotlin
HarmonEyesSDK.addGazeSamples(handle, samples)
```

`stopSession(handle)` — Stop and flush the session
- `handle: Long` — session handle e.g. `140532891648`
- returns `Boolean` — true on success e.g. `true`
```kotlin
HarmonEyesSDK.stopSession(handle)
```

`destroy(handle)` — Release native resources
- `handle: Long` — session handle e.g. `140532891648`
```kotlin
HarmonEyesSDK.destroy(handle)
```

`getCogLoad(handle)` — Poll latest cognitive load result
- `handle: Long` — session handle e.g. `140532891648`
- returns `CognitiveLoadPrediction?` — null if no result yet e.g. `CognitiveLoadPrediction(High, 0.87, 0.92, 5)`
```kotlin
val cogLoad = HarmonEyesSDK.getCogLoad(handle)
```

`getSleepiness(handle)` — Poll latest sleepiness result
- `handle: Long` — session handle e.g. `140532891648`
- returns `SleepinessResult?` — null if no result yet e.g. `SleepinessResult(Alert, 0.91, 3)`
```kotlin
val sleepiness = HarmonEyesSDK.getSleepiness(handle)
```

`isRunning(handle)` — Whether session is active
- `handle: Long` — session handle e.g. `140532891648`
- returns `Boolean` e.g. `true`
```kotlin
val active = HarmonEyesSDK.isRunning(handle)
```

`batchCount(handle)` — Number of batches processed so far
- `handle: Long` — session handle e.g. `140532891648`
- returns `Int` e.g. `12`
```kotlin
val count = HarmonEyesSDK.batchCount(handle)
```

`licenseInfo(handle)` — License status and expiry
- `handle: Long` — session handle e.g. `140532891648`
- returns `LicenseInfo?` e.g. `LicenseInfo(status="active", expiryDate="2026-12-31")`
```kotlin
val license = HarmonEyesSDK.licenseInfo(handle)
```

`getSdkVersionString()` — SDK version
- returns `String` e.g. `"1.0.0"`
```kotlin
val version = HarmonEyesSDK.getSdkVersionString()
```

`getFeatureVersionString()` — Feature model version
- returns `String` e.g. `"1.0.0"`
```kotlin
val features = HarmonEyesSDK.getFeatureVersionString()
```

## Pupil Labs Neon Integration

The SDK provides native integration with the Pupil Labs Neon eye tracker, including device discovery, real-time data streaming, and recording support.

### Permissions

If using Pupil Labs Neon, add these additional permissions:

```xml
<uses-permission android:name="android.permission.ACCESS_WIFI_STATE" />
<uses-permission android:name="android.permission.CHANGE_WIFI_MULTICAST_STATE" />
```

- `ACCESS_WIFI_STATE` / `CHANGE_WIFI_MULTICAST_STATE` — required for mDNS device discovery

### Neon API

`neonCreate()` — Create a Neon client with default options
- returns `Long` — Neon client handle e.g. `140532891648`
```kotlin
val neonHandle = HarmonEyesSDK.neonCreate()
```

`neonCreateWithOptions(options)` — Create a Neon client with custom options
- `options: NeonOptions` — discovery timeout, buffer size, sample rate e.g. `NeonOptions(sampleRateHz = 200)`
- returns `Long` — Neon client handle e.g. `140532891648`
```kotlin
val neonHandle = HarmonEyesSDK.neonCreateWithOptions(NeonOptions(sampleRateHz = 200))
```

`neonConnectDevice(handle, device)` — Connect to a discovered Neon device
- `handle: Long` — Neon client handle e.g. `140532891648`
- `device: NeonDevice` — device from mDNS discovery e.g. `NeonDevice(ip="192.168.1.42", port=8080)`
- returns `Boolean` — true on success e.g. `true`
```kotlin
HarmonEyesSDK.neonConnectDevice(neonHandle, device)
```

`neonStartDataStreaming(handle, callback)` — Start streaming preprocessed eye data
- `handle: Long` — Neon client handle e.g. `140532891648`
- `callback: PreprocessedSampleListener` — called for each sample
- returns `Boolean` — true on success e.g. `true`
```kotlin
HarmonEyesSDK.neonStartDataStreaming(neonHandle) { sample -> ... }
```

`neonStopDataStreaming(handle)` — Stop streaming
- `handle: Long` — Neon client handle e.g. `140532891648`
- returns `Boolean` — true on success e.g. `true`
```kotlin
HarmonEyesSDK.neonStopDataStreaming(neonHandle)
```

`neonDisconnect(handle)` — Disconnect from device
- `handle: Long` — Neon client handle e.g. `140532891648`
- returns `Boolean` — true on success e.g. `true`
```kotlin
HarmonEyesSDK.neonDisconnect(neonHandle)
```

`neonDestroy(handle)` — Release native resources
- `handle: Long` — Neon client handle e.g. `140532891648`
```kotlin
HarmonEyesSDK.neonDestroy(neonHandle)
```

`neonIsConnected(handle)` — Connection status
- `handle: Long` — Neon client handle e.g. `140532891648`
- returns `Boolean` e.g. `true`
```kotlin
val connected = HarmonEyesSDK.neonIsConnected(neonHandle)
```

`neonIsDataStreaming(handle)` — Whether data streaming is active
- `handle: Long` — Neon client handle e.g. `140532891648`
- returns `Boolean` e.g. `true`
```kotlin
val streaming = HarmonEyesSDK.neonIsDataStreaming(neonHandle)
```

`neonGetCalibration(handle)` — Camera calibration data
- `handle: Long` — Neon client handle e.g. `140532891648`
- returns `CameraCalibration` e.g. `CameraCalibration(fx=883.4, fy=883.4, cx=640.0, cy=400.0)`
```kotlin
val calibration = HarmonEyesSDK.neonGetCalibration(neonHandle)
```

`neonStartRecording(handle)` — Start a recording on the device
- `handle: Long` — Neon client handle e.g. `140532891648`
- returns `String?` — recording ID e.g. `"abc123-recording"`
```kotlin
val recordingId = HarmonEyesSDK.neonStartRecording(neonHandle)
```

`neonStopRecording(handle)` — Stop the current recording
- `handle: Long` — Neon client handle e.g. `140532891648`
- returns `Boolean` — true on success e.g. `true`
```kotlin
HarmonEyesSDK.neonStopRecording(neonHandle)
```

`neonCancelRecording(handle)` — Cancel and discard the current recording
- `handle: Long` — Neon client handle e.g. `140532891648`
- returns `Boolean` — true on success e.g. `true`
```kotlin
HarmonEyesSDK.neonCancelRecording(neonHandle)
```

`neonSendEvent(handle, name, timestampNs)` — Send a named event marker to the recording
- `handle: Long` — Neon client handle from `neonCreate()` e.g. `140532891648`
- `name: String` — event label e.g. `"task-start"`
- `timestampNs: Long` — event timestamp in nanoseconds e.g. `1712695483917000000`
- returns `Boolean` — true on success e.g. `true`
```kotlin
HarmonEyesSDK.neonSendEvent(neonHandle, "task-start", System.nanoTime())
```

`neonDeviceInfo(handle)` — Get connected device info
- `handle: Long` — Neon client handle e.g. `140532891648`
- returns `NeonDevice?` — null if not connected e.g. `NeonDevice(name="Neon Sensor Module", ip="192.168.1.42", port=8080)`
```kotlin
val device = HarmonEyesSDK.neonDeviceInfo(neonHandle)
```

## Architecture

- **ARM64 only** — the native engine (`libtheia.so`) is compiled for `arm64-v8a`
- **Two native libraries** ship in the AAR:
  - `libtheia.so` — vendor C engine (~22MB)
  - `libharmoneyes_jni.so` — JNI bridge
- **Handle-based JNI** — each session/client holds its own native pointer, no global singletons
- **Handle lifecycle** — call `destroy(handle)` / `neonDestroy(handle)` when done to release native resources
- **Thread-safe callbacks** — C library callbacks route through JNI to your Kotlin listeners on a background thread

## License

This SDK is distributed under a commercial license. A valid license key is required for runtime usage. See [license.md](license.md) for details.
