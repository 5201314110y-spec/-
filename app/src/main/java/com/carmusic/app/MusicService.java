package com.carmusic.app;

import android.app.Service;
import android.content.Intent;
import android.os.Bundle;
import android.os.IBinder;
import android.support.v4.media.MediaBrowserCompat;
import android.support.v4.media.MediaMetadataCompat;
import android.support.v4.media.session.MediaSessionCompat;
import android.support.v4.media.session.PlaybackStateCompat;

import androidx.annotation.NonNull;
import androidx.annotation.Nullable;
import androidx.media.MediaBrowserServiceCompat;

import java.lang.ref.WeakReference;
import java.util.ArrayList;
import java.util.List;

/**
 * MediaBrowserService - 让车机语音助手能发现并控制本应用
 *
 * 支持的语音指令：
 * - "播放音乐" → onPlay
 * - "暂停"     → onPause
 * - "下一首"   → onSkipToNext
 * - "上一首"   → onSkipToPrevious
 * - "播放周杰伦的晴天" → onPlayFromSearch
 *
 * 优化 4：使用 WeakReference 持有实例，避免 Service 销毁后静态强引用
 * 阻止 GC 回收，消除内存泄漏风险。外部统一通过 getInstance() 访问，
 * 返回 null 时表示 Service 已销毁，调用方应做 null 检查。
 */
public class MusicService extends MediaBrowserServiceCompat {

    private MediaSessionCompat mediaSession;
    private PlaybackStateCompat.Builder stateBuilder;

    // WeakReference：Service 销毁后 GC 可正常回收，不会因静态强引用泄漏
    private static WeakReference<MusicService> instanceRef;

    /** 获取当前 Service 实例，Service 已销毁时返回 null */
    public static MusicService getInstance() {
        return instanceRef != null ? instanceRef.get() : null;
    }

    private WebBridge webBridge;

    @Override
    public void onCreate() {
        super.onCreate();
        instanceRef = new WeakReference<>(this);

        // 创建 MediaSession
        mediaSession = new MediaSessionCompat(this, "CarMusicSession");

        // 设置支持的操作
        stateBuilder = new PlaybackStateCompat.Builder()
            .setActions(
                PlaybackStateCompat.ACTION_PLAY |
                PlaybackStateCompat.ACTION_PAUSE |
                PlaybackStateCompat.ACTION_PLAY_PAUSE |
                PlaybackStateCompat.ACTION_SKIP_TO_NEXT |
                PlaybackStateCompat.ACTION_SKIP_TO_PREVIOUS |
                PlaybackStateCompat.ACTION_STOP |
                PlaybackStateCompat.ACTION_SEEK_TO |
                PlaybackStateCompat.ACTION_PLAY_FROM_SEARCH
            );

        mediaSession.setPlaybackState(stateBuilder.build());
        mediaSession.setCallback(new MediaSessionCallback());
        mediaSession.setActive(true);

        setSessionToken(mediaSession.getSessionToken());
    }

    public MediaSessionCompat getMediaSession() {
        return mediaSession;
    }

    public void setWebBridge(WebBridge bridge) {
        this.webBridge = bridge;
    }

    /**
     * 更新正在播放的元数据（从 WebView 回调）
     */
    public void updateMetadata(String title, String artist, String album, long duration) {
        MediaMetadataCompat metadata = new MediaMetadataCompat.Builder()
            .putString(MediaMetadataCompat.METADATA_KEY_TITLE, title)
            .putString(MediaMetadataCompat.METADATA_KEY_ARTIST, artist)
            .putString(MediaMetadataCompat.METADATA_KEY_ALBUM, album)
            .putLong(MediaMetadataCompat.METADATA_KEY_DURATION, duration)
            .build();
        mediaSession.setMetadata(metadata);
    }

    /**
     * 更新播放状态
     */
    public void updatePlaybackState(boolean isPlaying, long position) {
        int state = isPlaying ? PlaybackStateCompat.STATE_PLAYING : PlaybackStateCompat.STATE_PAUSED;
        mediaSession.setPlaybackState(
            stateBuilder.setState(state, position, 1.0f).build()
        );
    }

    @Nullable
    @Override
    public BrowserRoot onGetRoot(@NonNull String clientPackageName, int clientUid, @Nullable Bundle rootHints) {
        // 允许所有客户端连接（车机语音助手）
        return new BrowserRoot("root", null);
    }

    @Override
    public void onLoadChildren(@NonNull String parentId, @NonNull Result<List<MediaBrowserCompat.MediaItem>> result) {
        // 返回空列表（我们主要通过搜索播放）
        result.sendResult(new ArrayList<>());
    }

    /**
     * MediaSession 回调 - 处理语音指令
     */
    private class MediaSessionCallback extends MediaSessionCompat.Callback {

        @Override
        public void onPlay() {
            if (webBridge != null) {
                webBridge.executeCommand("togglePlay", "");
            }
        }

        @Override
        public void onPause() {
            if (webBridge != null) {
                webBridge.executeCommand("pause", "");
            }
        }

        @Override
        public void onSkipToNext() {
            if (webBridge != null) {
                webBridge.executeCommand("next", "");
            }
        }

        @Override
        public void onSkipToPrevious() {
            if (webBridge != null) {
                webBridge.executeCommand("prev", "");
            }
        }

        @Override
        public void onStop() {
            if (webBridge != null) {
                webBridge.executeCommand("stop", "");
            }
        }

        @Override
        public void onSeekTo(long pos) {
            if (webBridge != null) {
                webBridge.executeCommand("seekTo", String.valueOf(pos));
            }
        }

        @Override
        public void onPlayFromSearch(String query, Bundle extras) {
            // 语音搜索播放："播放XXX"
            if (webBridge != null) {
                webBridge.executeCommand("voiceSearch", query != null ? query : "");
            }
        }
    }

    @Override
    public void onDestroy() {
        if (mediaSession != null) {
            mediaSession.setActive(false);
            mediaSession.release();
        }
        instanceRef = null;
        super.onDestroy();
    }
}
