# Streamlit课表生成器

## 功能介绍

本应用程序可以将Excel格式的课程表转换为两种格式：
1. **教室课表** - 适合教师和管理员查看的详细课表
2. **教室班牌** - 适合在教室门口显示的简洁课表

## 使用方法

1. 安装依赖：
   ```
   pip install -r requirements.txt
   ```

2. 运行应用：
   ```
   streamlit run app.py
   ```

3. 在浏览器中打开应用
   - 上传Excel源文件
   - 选择转换类型
   - 输入周数
   - 点击生成按钮
   - 下载生成的课表文件

## 文件说明

- `app.py` - Streamlit主应用
- `utils.py` - 课表生成逻辑
- `requirements.txt` - Python依赖包
- `README.md` - 使用说明

## 注意事项

- 源文件必须是Excel格式（.xlsx或.xls）
- 源文件的格式应与系统要求的格式一致
