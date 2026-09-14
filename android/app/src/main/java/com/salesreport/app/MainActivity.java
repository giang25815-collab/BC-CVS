package com.salesreport.app;

import android.app.Activity;
import android.content.ClipData;
import android.content.ContentValues;
import android.content.Intent;
import android.net.Uri;
import android.os.Build;
import android.os.Bundle;
import android.os.Environment;
import android.provider.MediaStore;
import android.util.Base64;
import android.util.Log;
import android.webkit.JavascriptInterface;
import android.webkit.ValueCallback;
import android.webkit.WebChromeClient;
import android.webkit.WebResourceRequest;
import android.webkit.WebResourceResponse;
import android.webkit.WebSettings;
import android.webkit.WebView;
import android.widget.Toast;

import androidx.activity.OnBackPressedCallback;
import androidx.activity.result.ActivityResultLauncher;
import androidx.activity.result.contract.ActivityResultContracts;
import androidx.annotation.NonNull;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.FileProvider;
import androidx.webkit.WebViewAssetLoader;
import androidx.webkit.WebViewClientCompat;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

public class MainActivity extends AppCompatActivity {

    private static final String TAG = "SalesReportApp";
    private WebView mWebView;
    private ValueCallback<Uri[]> mFilePathCallback;
    private final ExecutorService mExecutor = Executors.newSingleThreadExecutor();

    private final ActivityResultLauncher<Intent> mFileChooserLauncher =
            registerForActivityResult(new ActivityResultContracts.StartActivityForResult(), result -> {
                if (mFilePathCallback == null) return;

                Uri[] results = null;
                if (result.getResultCode() == Activity.RESULT_OK && result.getData() != null) {
                    Intent data = result.getData();
                    ClipData clipData = data.getClipData();
                    if (clipData != null) {
                        results = new Uri[clipData.getItemCount()];
                        for (int i = 0; i < clipData.getItemCount(); i++) {
                            results[i] = clipData.getItemAt(i).getUri();
                        }
                    } else if (data.getData() != null) {
                        results = new Uri[]{data.getData()};
                    }
                }

                mFilePathCallback.onReceiveValue(results);
                mFilePathCallback = null;
            });

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        mWebView = findViewById(R.id.webview);
        setupWebView();

        // Handle native back press
        getOnBackPressedDispatcher().addCallback(this, new OnBackPressedCallback(true) {
            @Override
            public void handleOnBackPressed() {
                if (mWebView != null && mWebView.canGoBack()) {
                    mWebView.goBack();
                } else {
                    finish();
                }
            }
        });
    }

    private void setupWebView() {
        WebSettings settings = mWebView.getSettings();
        settings.setJavaScriptEnabled(true);
        settings.setDomStorageEnabled(true);
        settings.setDatabaseEnabled(true);
        settings.setAllowFileAccess(true);
        settings.setAllowContentAccess(true);
        settings.setUseWideViewPort(true);
        settings.setLoadWithOverviewMode(true);
        settings.setSupportZoom(false);
        settings.setBuiltInZoomControls(false);
        settings.setDisplayZoomControls(false);

        // Hardware acceleration
        mWebView.setLayerType(WebView.LAYER_TYPE_HARDWARE, null);

        // Setup AssetLoader for secure HTTPS origin
        final WebViewAssetLoader assetLoader = new WebViewAssetLoader.Builder()
                .addPathHandler("/assets/", new WebViewAssetLoader.AssetsPathHandler(this))
                .build();

        mWebView.setWebViewClient(new WebViewClientCompat() {
            @Override
            public WebResourceResponse shouldInterceptRequest(WebView view, WebResourceRequest request) {
                return assetLoader.shouldInterceptRequest(request.getUrl());
            }
        });

        // Setup file chooser & bridge
        mWebView.setWebChromeClient(new WebChromeClient() {
            @Override
            public boolean onShowFileChooser(WebView webView, ValueCallback<Uri[]> filePathCallback, FileChooserParams fileChooserParams) {
                if (mFilePathCallback != null) {
                    mFilePathCallback.onReceiveValue(null);
                }
                mFilePathCallback = filePathCallback;

                // Luôn mở bộ chọn tệp phổ quát (*/*) để Android không làm mờ file .xlsb (do Android thiếu định nghĩa MIME cho .xlsb)
                Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
                intent.addCategory(Intent.CATEGORY_OPENABLE);
                intent.setType("*/*");
                if (fileChooserParams != null && fileChooserParams.getMode() == FileChooserParams.MODE_OPEN_MULTIPLE) {
                    intent.putExtra(Intent.EXTRA_ALLOW_MULTIPLE, true);
                }

                try {
                    mFileChooserLauncher.launch(intent);
                } catch (Exception e) {
                    Log.e(TAG, "Cannot launch file chooser", e);
                    mFilePathCallback.onReceiveValue(null);
                    mFilePathCallback = null;
                    return false;
                }
                return true;
            }
        });

        // Register AndroidBridge
        mWebView.addJavascriptInterface(new AndroidBridge(), "AndroidBridge");

        // Load entry page via HTTPS asset loader
        mWebView.loadUrl("https://appassets.androidplatform.net/assets/www/index.html");
    }

    public class AndroidBridge {

        @JavascriptInterface
        public void saveAndShare(final String base64Data, final String filename, final String mimeType) {
            mExecutor.execute(() -> {
                try {
                    byte[] fileBytes = Base64.decode(base64Data, Base64.DEFAULT);

                    // 1. Save to public Downloads directory
                    if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                        ContentValues values = new ContentValues();
                        values.put(MediaStore.MediaColumns.DISPLAY_NAME, filename);
                        values.put(MediaStore.MediaColumns.MIME_TYPE, mimeType);
                        values.put(MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS + "/BaoCaoDoanhSo");
                        Uri uri = getContentResolver().insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values);
                        if (uri != null) {
                            try (OutputStream os = getContentResolver().openOutputStream(uri)) {
                                if (os != null) {
                                    os.write(fileBytes);
                                }
                            }
                        }
                    } else {
                        File dir = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), "BaoCaoDoanhSo");
                        if (!dir.exists()) dir.mkdirs();
                        File dest = new File(dir, filename);
                        try (FileOutputStream fos = new FileOutputStream(dest)) {
                            fos.write(fileBytes);
                        }
                    }

                    // 2. Save to cache dir for sharing via FileProvider
                    File shareDir = new File(getCacheDir(), "shared_reports");
                    if (!shareDir.exists()) shareDir.mkdirs();
                    File shareFile = new File(shareDir, filename);
                    try (FileOutputStream fos = new FileOutputStream(shareFile)) {
                        fos.write(fileBytes);
                    }

                    Uri contentUri = FileProvider.getUriForFile(MainActivity.this, getPackageName() + ".provider", shareFile);

                    // 3. UI feedback & Share Dialog
                    runOnUiThread(() -> {
                        Toast.makeText(MainActivity.this, "Đã lưu vào Tải về: " + filename, Toast.LENGTH_SHORT).show();

                        Intent shareIntent = new Intent(Intent.ACTION_SEND);
                        shareIntent.setType(mimeType);
                        shareIntent.putExtra(Intent.EXTRA_STREAM, contentUri);
                        shareIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);
                        startActivity(Intent.createChooser(shareIntent, "Chia sẻ báo cáo: " + filename));
                    });

                } catch (Exception e) {
                    Log.e(TAG, "Error in saveAndShare", e);
                    runOnUiThread(() -> Toast.makeText(MainActivity.this, "Lỗi lưu file: " + e.getMessage(), Toast.LENGTH_LONG).show());
                }
            });
        }

        @JavascriptInterface
        public void shareToZalo(final String base64Data, final String filename) {
            mExecutor.execute(() -> {
                try {
                    byte[] fileBytes = Base64.decode(base64Data, Base64.DEFAULT);

                    // 1. Save to Downloads
                    if (Build.VERSION.SDK_INT >= Build.VERSION_CODES.Q) {
                        ContentValues values = new ContentValues();
                        values.put(MediaStore.MediaColumns.DISPLAY_NAME, filename);
                        values.put(MediaStore.MediaColumns.MIME_TYPE, "image/png");
                        values.put(MediaStore.MediaColumns.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS + "/BaoCaoDoanhSo");
                        Uri uri = getContentResolver().insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values);
                        if (uri != null) {
                            try (OutputStream os = getContentResolver().openOutputStream(uri)) {
                                if (os != null) {
                                    os.write(fileBytes);
                                }
                            }
                        }
                    } else {
                        File dir = new File(Environment.getExternalStoragePublicDirectory(Environment.DIRECTORY_DOWNLOADS), "BaoCaoDoanhSo");
                        if (!dir.exists()) dir.mkdirs();
                        File dest = new File(dir, filename);
                        try (FileOutputStream fos = new FileOutputStream(dest)) {
                            fos.write(fileBytes);
                        }
                    }

                    // 2. Save to cache dir for sharing via FileProvider
                    File shareDir = new File(getCacheDir(), "shared_reports");
                    if (!shareDir.exists()) shareDir.mkdirs();
                    File shareFile = new File(shareDir, filename);
                    try (FileOutputStream fos = new FileOutputStream(shareFile)) {
                        fos.write(fileBytes);
                    }

                    Uri contentUri = FileProvider.getUriForFile(MainActivity.this, getPackageName() + ".provider", shareFile);

                    // 3. UI feedback & direct Zalo or Share Chooser
                    runOnUiThread(() -> {
                        Toast.makeText(MainActivity.this, "Đang mở Zalo gửi báo cáo...", Toast.LENGTH_SHORT).show();

                        Intent shareIntent = new Intent(Intent.ACTION_SEND);
                        shareIntent.setType("image/png");
                        shareIntent.putExtra(Intent.EXTRA_STREAM, contentUri);
                        shareIntent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION);

                        // Cố gắng mở trực tiếp Zalo
                        Intent zaloIntent = new Intent(shareIntent);
                        zaloIntent.setPackage("com.zing.zalo");

                        try {
                            if (getPackageManager().getLaunchIntentForPackage("com.zing.zalo") != null) {
                                startActivity(zaloIntent);
                                return;
                            }
                        } catch (Exception ignored) { }

                        // Nếu không có Zalo hoặc lỗi, mở bộ chọn ứng dụng
                        startActivity(Intent.createChooser(shareIntent, "Gửi báo cáo qua Zalo / Tin nhắn"));
                    });

                } catch (Exception e) {
                    Log.e(TAG, "Error in shareToZalo", e);
                    runOnUiThread(() -> Toast.makeText(MainActivity.this, "Lỗi gửi Zalo: " + e.getMessage(), Toast.LENGTH_LONG).show());
                }
            });
        }

        @JavascriptInterface
        public void copyToClipboard(final String text) {
            runOnUiThread(() -> {
                try {
                    android.content.ClipboardManager clipboard = (android.content.ClipboardManager) getSystemService(CLIPBOARD_SERVICE);
                    android.content.ClipData clip = android.content.ClipData.newPlainText("Link Báo Cáo", text);
                    if (clipboard != null) {
                        clipboard.setPrimaryClip(clip);
                        Toast.makeText(MainActivity.this, "✅ Đã sao chép link báo cáo!", Toast.LENGTH_SHORT).show();
                    }
                } catch (Exception e) {
                    Log.e(TAG, "copyToClipboard error", e);
                }
            });
        }

        @JavascriptInterface
        public void shareText(final String text, final String title) {
            runOnUiThread(() -> {
                try {
                    Intent sendIntent = new Intent();
                    sendIntent.setAction(Intent.ACTION_SEND);
                    sendIntent.putExtra(Intent.EXTRA_TEXT, text);
                    sendIntent.setType("text/plain");
                    Intent shareIntent = Intent.createChooser(sendIntent, title != null ? title : "Chia sẻ link báo cáo");
                    startActivity(shareIntent);
                } catch (Exception e) {
                    Log.e(TAG, "shareText error", e);
                }
            });
        }
    }

    @Override
    protected void onDestroy() {
        super.onDestroy();
        mExecutor.shutdown();
        if (mWebView != null) {
            mWebView.destroy();
        }
    }
}
