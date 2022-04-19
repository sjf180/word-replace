package com.aslack.w2r.word2replace;

import com.aspose.words.*;
import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.ApplicationRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Objects;
import java.util.stream.Stream;

@SpringBootApplication
public class Word2ReplaceApplication implements ApplicationRunner {

	public static String readPath;
	public static String savePath;
	public static String pic1;

	public static void main(String[] args) {
		readPath = args[0];
		savePath = args[1];
		pic1 = args[2];
		System.out.println("读取目录：" + readPath + ", 存储目录：" + savePath);
		SpringApplication.run(Word2ReplaceApplication.class, args);
	}

	@Override
	public void run(ApplicationArguments args) throws Exception {
		String dir = System.getProperty("user.dir");
		System.out.println("工作目录：" + dir);
		try(Stream<Path> paths = Files.walk(Paths.get(readPath))) {
			paths.filter(Files::isRegularFile).forEach(v -> {
				if (!v.startsWith("~$") && (v.toString().toLowerCase().endsWith(".docx") || v.toString().toLowerCase().endsWith(".doc"))) {
					try {
						System.out.println("正在处理：" + v);
						Document doc = new Document(new FileInputStream(v.toFile()));
						DocumentBuilder builder = new DocumentBuilder(doc);
						NodeCollection nodeCollection = doc.getChildNodes(NodeType.RUN, true);

						for (int i = 0; i < nodeCollection.getCount(); i++) {
							Node node = nodeCollection.get(i);
							if (node.getNodeType() == NodeType.RUN) {
								Run run = (Run) node;
								Color color = run.getFont().getColor();
								if (Objects.equals(color, new Color(255, 0, 0))) {

									boolean flag = false;
									if (null != pic1) {
										String[] s = pic1.split(",");
										for (int j = 0; j < s.length; j++) {
											String n = s[j].substring(0, s[j].lastIndexOf("."));
											if (run.getText().contains(n)) {
												node.getRange().replace(run.getText(), " ");
												builder.moveTo(run);
												BufferedImage bf = ImageIO.read(new File(dir + "\\pic\\" + s[j]));
												builder.insertImage(dir + "\\pic\\" + s[j], RelativeHorizontalPosition.PAGE, 0, RelativeVerticalPosition.PAGE, 0, bf.getWidth(), bf.getHeight(), WrapType.INLINE);
												bf.flush();
												flag = true;
											}
										}
									}

									if (!flag) {
										int len = run.getText().length();
										StringBuilder s = new StringBuilder();
										for (int j = 0; j < len; j++) {
											s.append(" ");
										}
										node.getRange().replace(run.getText(), s.toString());
									}
								}
							}
						}

						String s = v.toString().replace(readPath, savePath);
						doc.save(s);
						doc.cleanup();
						System.out.println("处理完成：" + s);
					} catch (Exception e) {
						e.printStackTrace();
					}
				}
			});
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
