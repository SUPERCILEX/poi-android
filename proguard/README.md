# POI Android proguard

This module is just a wrapper around a proguard config. We can't directly include POI and the
proguard config together because the shadow transpiler needs the Gradle `java` plugin to work which
doesn't mix well with the Android Gradle Plugin.

This config is highly optimized for [Robot Scouter](https://github.com/SUPERCILEX/Robot-Scouter). If
you need a different configuration, please fork this repo and use your own custom build.
