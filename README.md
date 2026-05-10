# 🚗 车载音乐 APK

竖屏车机适配的聚合音乐播放器 Android 应用。

## 项目结构

```
.
├── build.gradle                    # 根构建文件
├── settings.gradle
├── gradlew / gradlew.bat           # Gradle Wrapper
├── gradle/wrapper/
│   ├── gradle-wrapper.jar
│   └── gradle-wrapper.properties
├── .github/workflows/build.yml     # CI 自动编译 + Release
├── app/
│   ├── build.gradle                # 应用构建配置
│   ├── proguard-rules.pro
│   └── src/main/
│       ├── AndroidManifest.xml
│       ├── java/com/carmusic/app/
│       │   ├── MainActivity.java
│       │   ├── MusicService.java
│       │   └── WebBridge.java
│       ├── res/mipmap-*/
│       │   └── ic_launcher.png     # 各密度图标
│       ├── res/xml/
│       │   └── network_security_config.xml
│       └── assets/www/
│           └── index.html          # 前端页面
```

## 编译方法

### 方法一：Android Studio（推荐）

1. 用 Android Studio 打开仓库根目录
2. 等待 Gradle 同步完成
3. Build → Build Bundle(s) / APK(s) → Build APK(s)
4. APK 输出在 `app/build/outputs/apk/debug/app-debug.apk`

### 方法二：命令行编译

```bash
chmod +x gradlew
./gradlew assembleDebug

# 或使用系统 gradle
gradle assembleDebug
```

### 方法三：在线编译（无需本地环境，已配置）

本仓库已内置 GitHub Actions 工作流 `.github/workflows/build.yml`：

- 任意分支 push / PR / 手动触发 → 自动编译 Debug APK，作为 Actions artifact 下载
- 推送 `v*` tag（如 `v1.0.0`）→ 自动编译并创建 GitHub Release，附带 APK

发布新版本：

```bash
git tag v1.0.0
git push origin v1.0.0
```

随后在仓库的 Releases 页面即可下载 `car-music-v1.0.0-debug.apk`。

## 安装到车机

```bash
# USB 连接车机后
adb install app-debug.apk

# 或拷贝到U盘，车机文件管理器安装
```

## 功能

- 聚合搜索：网易云、QQ、酷狗、酷我、咪咕、B站
- 竖屏全屏：适配竖屏车机，沉浸式无状态栏
- 大按钮交互：驾驶时易操作
- 歌词同步：全屏歌词浮层
- 自动连播：播完自动下一首
- 常亮屏幕：播放时不息屏

## 注意事项

- 需要网络连接（WiFi 或车载流量）
- 音频 URL 有时效性，断网后需重新搜索
- 部分歌曲因版权限制可能无法播放
- 最低支持 Android 5.0 (API 21)
