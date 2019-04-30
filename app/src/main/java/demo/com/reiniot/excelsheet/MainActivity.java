package demo.com.reiniot.excelsheet;

import android.app.Activity;
import android.content.pm.PackageManager;
import android.os.Bundle;
import android.support.annotation.NonNull;
import android.support.v4.app.ActivityCompat;
import android.support.v7.app.AppCompatActivity;
import android.view.Menu;
import android.view.MenuItem;
import android.view.View;

import demo.com.reiniot.lib.ExcelManager;

public class MainActivity extends AppCompatActivity {

    private static String[] PERMISSIONS_STORAGE = {
            "android.permission.READ_EXTERNAL_STORAGE",
            "android.permission.WRITE_EXTERNAL_STORAGE" };
    private static final int REQUEST_EXTERNAL_STORAGE = 1;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        findViewById(R.id.btn).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View view) {
                excelOperation();
            }
        });

        verifyStoragePermissions(this);

    }

    private void excelOperation(){
        new Thread(new Runnable() {
            @Override
            public void run() {
                ExcelManager manager = ExcelManager.getSingleInstance();
                //设置大文件路径
                //设置大文件中和小文件中以某个字段为对比
                //设置小文件路径
                //设置大文件中和小文件中以某个字段为对比
                //比如
                //大excel文件有    身高  体重   三维
                //小exclel文件中有    三维  性别   体重
                //需要以体重为指标挑选大文件和小文件中共同的部分  大文件 dentifyId = 1   小文件dentifyId = 2

                manager.setBigFilePath("/sdcard/mobile_all.xlsx").setBigDentifyId(3).setSmallFilePath("/sdcard/mobile.xlsx").setSmallDentifyId(0);
                manager.setRootPath("/sdcard/");
                manager.excute();
            }
        }).start();
    }


    public static void verifyStoragePermissions(Activity activity) {
        try {
            //检测是否有写的权限
            int permission = ActivityCompat.checkSelfPermission(activity,
                    "android.permission.WRITE_EXTERNAL_STORAGE");
            if (permission != PackageManager.PERMISSION_GRANTED) {
                // 没有写的权限，去申请写的权限，会弹出对话框
                ActivityCompat.requestPermissions(activity, PERMISSIONS_STORAGE,REQUEST_EXTERNAL_STORAGE);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    @Override
    public boolean onCreateOptionsMenu(Menu menu) {
        // Inflate the menu; this adds items to the action bar if it is present.
        getMenuInflater().inflate(R.menu.menu_main, menu);
        return true;
    }

    @Override
    public boolean onOptionsItemSelected(MenuItem item) {
        // Handle action bar item clicks here. The action bar will
        // automatically handle clicks on the Home/Up button, so long
        // as you specify a parent activity in AndroidManifest.xml.
        int id = item.getItemId();

        //noinspection SimplifiableIfStatement
        if (id == R.id.action_settings) {
            return true;
        }

        return super.onOptionsItemSelected(item);
    }
}
