package mypoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.InetSocketAddress;
import java.nio.ByteBuffer;
import java.nio.channels.FileChannel;
import java.nio.channels.SelectionKey;
import java.nio.channels.Selector;
import java.nio.channels.ServerSocketChannel;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.locks.ReentrantLock;


public class LargetData {


    public static void main(String[] args) throws Exception {
        writeFile();
    }


    static void test() throws IOException {

        Path path2 = Paths.get("C:\\Users\\Administrator\\Desktop\\222222222.txt");
        if (Files.notExists(path2)) {
            Files.createFile(path2);
        } else {
            Files.delete(path2);
        }
    }


    static void readFile() throws Exception {
        try (FileInputStream inputStream = new FileInputStream(new File("C:\\Users\\Administrator\\Desktop\\flowdemo.txt"))) {
            FileChannel channel = inputStream.getChannel();
            ByteBuffer buffer = ByteBuffer.allocate(1024);
            channel.read(buffer);
        }
    }


    static void writeFile() throws Exception {
        try (FileOutputStream outputStream = new FileOutputStream("C:\\Users\\Administrator\\Desktop\\test.txt")) {
            FileChannel channel = outputStream.getChannel();
            int length = 1024;
            ByteBuffer buffer = ByteBuffer.allocate(length);
            String message = "hello world ...         1";
            buffer.put(message.getBytes());
            buffer.flip();
            channel.write(buffer);
        }
    }


    static void lockTest() {
        ReentrantLock lock = new ReentrantLock();
        lock.lock();
        lock.unlock();




    }


    private class ThredDemo implements Runnable {

        @Override
        public void run() {

        }
    }


    static void nioTest() throws Exception {
        ServerSocketChannel serverChannel = ServerSocketChannel.open();
        serverChannel.configureBlocking(false);
        serverChannel.socket().bind(new InetSocketAddress(1910));
        Selector selector = Selector.open();
        serverChannel.register(selector, SelectionKey.OP_ACCEPT);
        while (true) {
            int n = selector.select();
            if (n > 0) {
                Iterator itr = selector.selectedKeys().iterator();
                while (itr.hasNext()) {
                    SelectionKey key = (SelectionKey) itr.next();


                    itr.remove();
                }
            }
        }
    }


    void threadPoolTest() {
        ThreadPoolExecutor executor = new ThreadPoolExecutor(1, 1, 1l, TimeUnit.DAYS, null);

        executor.execute(new Runnable() {
            @Override
            public void run() {
                System.err.println("bbbb");
            }
        });

    }
}
