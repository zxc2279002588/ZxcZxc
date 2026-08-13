package cn.zhangruijunlaoshi.rundown;

import android.content.ContentValues;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.provider.MediaStore;
import android.util.Base64;
import android.webkit.JavascriptInterface;
import android.webkit.WebView;
import android.widget.Toast;

import com.getcapacitor.BridgeActivity;

import java.io.OutputStream;
import java.net.URLDecoder;
import java.nio.charset.StandardCharsets;

public class MainActivity extends BridgeActivity {
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        WebView webView = bridge.getWebView();
        webView.getSettings().setMediaPlaybackRequiresUserGesture(false);
        webView.addJavascriptInterface(new AndroidDownloads(), "AndroidDownloads");
        webView.setDownloadListener((url, userAgent, contentDisposition, mimeType, contentLength) -> {
            if (!url.startsWith("blob:")) return;
            String fileName = downloadName(contentDisposition);
            String script = "fetch(" + quote(url) + ").then(r=>r.blob()).then(b=>{" +
                "const reader=new FileReader();reader.onloadend=()=>AndroidDownloads.save(" +
                "reader.result," + quote(fileName) + "," + quote(mimeType) + ");reader.readAsDataURL(b);})";
            webView.evaluateJavascript(script, null);
        });
    }

    private static String quote(String value) {
        if (value == null) return "\"\"";
        return "\"" + value.replace("\\", "\\\\").replace("\"", "\\\"") + "\"";
    }

    private static String downloadName(String contentDisposition) {
        if (contentDisposition != null) {
            int index = contentDisposition.indexOf("filename=");
            if (index >= 0) {
                String name = contentDisposition.substring(index + 9).replace("\"", "").trim();
                try { return URLDecoder.decode(name, StandardCharsets.UTF_8.name()); }
                catch (Exception ignored) { return name; }
            }
        }
        return "新媒体串单导出.xlsx";
    }

    private class AndroidDownloads {
        @JavascriptInterface
        public void save(String dataUrl, String fileName, String mimeType) {
            try {
                byte[] bytes = Base64.decode(dataUrl.substring(dataUrl.indexOf(',') + 1), Base64.DEFAULT);
                ContentValues values = new ContentValues();
                values.put(MediaStore.Downloads.DISPLAY_NAME, fileName.endsWith(".xlsx") ? fileName : fileName + ".xlsx");
                values.put(MediaStore.Downloads.MIME_TYPE,
                    mimeType == null || mimeType.isEmpty()
                        ? "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        : mimeType);
                values.put(MediaStore.Downloads.RELATIVE_PATH, Environment.DIRECTORY_DOWNLOADS + "/新媒体串单");
                Uri destination = getContentResolver().insert(MediaStore.Downloads.EXTERNAL_CONTENT_URI, values);
                if (destination == null) throw new IllegalStateException("无法创建下载文件");
                try (OutputStream output = getContentResolver().openOutputStream(destination)) {
                    if (output == null) throw new IllegalStateException("无法写入下载文件");
                    output.write(bytes);
                }
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "已保存到 下载/新媒体串单", Toast.LENGTH_LONG).show());
            } catch (Exception error) {
                runOnUiThread(() -> Toast.makeText(MainActivity.this, "保存失败，请重试", Toast.LENGTH_LONG).show());
            }
        }
    }
}

