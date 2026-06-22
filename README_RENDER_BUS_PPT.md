# Render 部署说明：车体监测 PPT 自动生成

这个项目可以部署到 Render，给其他人通过网页上传 PPT 模版和照片压缩包，自动生成完整 PPT。

## 部署方式

1. 把当前项目推送到 GitHub。
2. 打开 Render，选择 New -> Blueprint。
3. 连接这个 GitHub 仓库。
4. Render 会读取 `render.yaml`，创建服务 `bus-ppt-generator`。
5. 部署完成后，打开 Render 给出的公开网址即可使用。

## Render 启动配置

`render.yaml` 已经配置为：

```bash
streamlit run app_bus_ppt_generator.py --server.port $PORT --server.address 0.0.0.0
```

## 使用说明

网页里需要上传：

- PPT 模版：`.pptx`
- 照片压缩包：`.zip` 或 `.rar`
- 客户、品牌、发布形式

照片压缩包建议结构：

```text
照片/
  6-0405-辽BM2910/
    1.jpg
    2.jpg
    3.jpg
    4.jpg
  10-6537-辽B09691D/
    1.jpg
    2.jpg
```

生成规则：

- 每个文件夹名写入“线路车牌”
- 每 2 张图片生成 1 页
- 同一个文件夹超过 2 张图片时，自动拆成连续多页
- 拆出的多页顶部表格内容完全一致
- 最终导出一个完整 `.pptx`

## 注意事项

- `.zip` 在 Render 上最稳定；`.rar` 依赖系统是否包含 `bsdtar`。
- 上传文件只在本次生成过程使用，不写入项目目录。
- 如果客户照片敏感，建议部署成私有服务或加访问密码。
