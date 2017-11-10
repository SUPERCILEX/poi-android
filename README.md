[![Build Status](https://travis-ci.org/SUPERCILEX/poi-android.svg?branch=master)](https://travis-ci.org/SUPERCILEX/poi-android)

# POI Android

POIA is a simple library enabling [Apache POI](https://poi.apache.org/) usage on Android.

## Table of Contents

- [Installation](#installation)
- [Updating Apache POI](#updating-apache-poi)
- [Notes](#notes)

## Installation

Add [JitPack](https://jitpack.io) to your `repositories`:

```groovy
allprojects {
    repositories {
        // ...
        maven { url 'https://jitpack.io' }
    }
}
```

And the POIA dependency itself:

```groovy
implementation "com.github.SUPERCILEX.poi-android:poi:$poiVersion"
```

If you're using proguard, also add:

```groovy
implementation "com.github.SUPERCILEX.poi-android:proguard:$poiVersion"
```

If you want source code and documentation, add the real Apache POI dependency as `compileOnly`:
```groovy
compileOnly "org.apache.poi:poi-ooxml:$poiVersion"
```

## Updating Apache POI

If you need a newer version of Apache POI than is provided by this transpiler, updating is as simple
as making a fork and changing a few lines of code:

1. Fork the repo and
   [update Apache POI](https://github.com/SUPERCILEX/poi-android/blob/0fceaa215ef5d752118a6768f4d436edf29b9b72/build.gradle#L15)
  1. PSA: you can find Apache POI release notes [here](https://poi.apache.org/changes.html)
1. Simply replace `SUPERCILEX` in the Gradle dependency with your own GitHub username
1. That's it, it's that simple! 🚀

## Notes

`XSSFWorkbook` (`*.xlsx`) does not work on pre-L (API < 21) devices. A simple solution is to show
the user some error message and gracefully downgrade to `HSSFWorkbook` (`*.xls`):

```kotlin
val workbook = if (isUnsupportedDevice) {
    showToast(getString(R.string.export_unsupported_device_rationale))
    HSSFWorkbook()
} else {
    XSSFWorkbook()
}

// Example unsupportedDevice property
val isUnsupportedDevice by lazy { VERSION.SDK_INT < VERSION_CODES.LOLLIPOP || isLowRamDevice }
```

Make sure to test your implementation thoroughly pre-L since `HSSFWorkbook` only supports a subset
of the `Workbook`'s APIs and might throw a UOE. Wikipedia even goes so far as to call it the
"Horrible SpreadSheet Format" so consider yourself warned. 😁
