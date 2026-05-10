package com.carmusic.app;

import android.os.Handler;
import android.os.Looper;
import android.util.Log;
import android.webkit.JavascriptInterface;
import android.webkit.WebView;

import org.json.JSONException;
import org.json.JSONObject;

/**
 * WebView ↔ Native 双向桥接
 *
 * JS → Native : window.CarMusic.updatePlayState / updateNowPlaying
 * Native → JS : webBridge.executeCommand(command, param)
 *
 * 安全说明（优化 3）：
 *   Native → JS 方向改用 JSONObject 序列化传参，彻底避免字符串拼接的 JS 注入风险。
 *   无论 param 包含引号、反斜杠、换行符或任意 Unicode，JSONObject.toString()
 *   均会正确转义；生成的 JSON 字面量由 JS 引擎解析而非执行，安全可靠。
 *
 *   同时移除旧的 escapeJs() 手工转义方法，统一由 JSON 序列化保证安全。
 */
public class WebBridge {

    private static final String TAG = "WebBridge";

    private final WebView webView;
    private final Handler mainHandler;

    public WebBridge(WebView webView) {
        this.webView = webView;
        this.mainHandler = new Handler(Looper.getMainLooper());
    }

    /**
     * Native → JS: 发送播放控制命令。
     *
     * 将 command 和 param 序列化为 JSON 对象字面量，嵌入自执行函数后
     * 通过 evaluateJavascript 执行。JS 侧读取 m.command / m.param，
     * 二者永远是字符串值，不会被解释为代码。
     *
     * JS 侧实现示例：
     *   window.nativeCommand = function(command, param) { ... }
     */
    public void executeCommand(final String command, final String param) {
        // 在调用线程预先序列化，避免占用主线程
        final String safeJson;
        try {
            JSONObject obj = new JSONObject();
            obj.put("command", command != null ? command : "");
            obj.put("param",   param   != null ? param   : "");
            safeJson = obj.toString(); // e.g. {"command":"next","param":""}
        } catch (JSONException e) {
            Log.e(TAG, "executeCommand: JSON 序列化失败 command=" + command, e);
            return;
        }

        mainHandler.post(new Runnable() {
            @Override
            public void run() {
                // safeJson 是合法 JSON 字面量，直接内联；
                // m.command / m.param 是字符串值，不会被当作代码执行。
                String js = "(function(){"
                        + "var m=" + safeJson + ";"
                        + "window.nativeCommand && window.nativeCommand(m.command,m.param);"
                        + "})();";
                webView.evaluateJavascript(js, null);
            }
        });
    }

    // =========================================================
    // JS → Native 接口（由 @JavascriptInterface 暴露给 WebView）
    // =========================================================

    /**
     * JS 通知 Native: 播放状态变化
     * 调用: window.CarMusic.updatePlayState(isPlaying, positionMs)
     */
    @JavascriptInterface
    public void updatePlayState(boolean isPlaying, long positionMs) {
        MusicService svc = MusicService.getInstance();
        if (svc != null) {
            svc.updatePlaybackState(isPlaying, positionMs);
        }
    }

    /**
     * JS 通知 Native: 当前歌曲信息变化
     * 调用: window.CarMusic.updateNowPlaying(title, artist, album, durationMs)
     */
    @JavascriptInterface
    public void updateNowPlaying(String title, String artist, String album, long durationMs) {
        MusicService svc = MusicService.getInstance();
        if (svc != null) {
            svc.updateMetadata(title, artist, album, durationMs);
        }
    }

    /**
     * JS 通知 Native: 播放列表已更新
     * 调用: window.CarMusic.onPlaylistUpdated(count)
     */
    @JavascriptInterface
    public void onPlaylistUpdated(int count) {
        // 可扩展：通知系统 MediaBrowser 刷新内容列表
    }
}
