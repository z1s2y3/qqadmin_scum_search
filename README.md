# SCUM物品查询插件 (AstrBot版本)

一个用于查询SCUM游戏物品代码的AstrBot插件，支持模糊搜索和官方中文名称翻译。

## 功能特性

- 🔍 **模糊搜索** - 支持按关键词模糊搜索物品
- 📦 **2247个物品** - 包含完整的SCUM物品数据库
- 🌐 **官方中文名称** - 自动翻译为游戏内官方中文名称
- 📋 **分类查询** - 支持按大分类和小分类浏览
- 🎮 **游戏命令** - 直接返回可复制的刷取命令

## 安装方法

### 步骤1: 复制插件文件

将整个 `qqadmin_scum_search` 文件夹复制到 AstrBot 的 `plugins` 目录：

```
你的AstrBot目录/
└── plugins/
    └── qqadmin_scum_search/
        ├── main.py
        ├── metadata.yaml
        ├── _conf_schema.json
        ├── __init__.py
        ├── requirements.txt
        ├── README.md
        └── SCUM全物品代码大全_豆包AI生成 (1).xlsx
```

### 步骤2: 安装依赖

```bash
pip install openpyxl>=3.0.0
```

### 步骤3: 重启AstrBot

重启 AstrBot 或在管理面板中刷新插件列表。

### 步骤4: 启用插件

在 AstrBot 管理面板中找到 "SCUM物品查询" 插件并启用它。

## 使用命令

### 基础查询
```
/scum <关键词>     - 模糊搜索物品
/scum分类          - 查看所有分类
/scum帮助          - 显示帮助菜单
```

### 分类查询
```
/scum武器          /scum弹药          /scum食物
/scum饮料          /scum医疗          /scum材料
/scum工具          /scum服装          /scum丧尸
/scum载具          /scum容器
```

### 模糊搜索示例
```
/scum 刀           # 搜索所有刀类武器
/scum ak47         # 搜索AK-47
/scum 医疗包        # 搜索医疗包
/scum zombie       # 搜索丧尸
/scum 摩托         # 搜索所有摩托车
```

## 配置选项

在 AstrBot 管理面板中可以配置：

| 配置项 | 类型 | 默认值 | 说明 |
|--------|------|--------|------|
| max_results | integer | 10 | 单次查询最大返回数量 |
| excel_path | string | SCUM全物品代码大全_豆包AI生成 (1).xlsx | Excel文件路径 |

## 游戏内使用

查询到物品代码后，在游戏控制台输入：
```
#SpawnItem 物品代码          # 刷出物品
#SpawnVehicle 载具代码       # 刷出载具
```

例如：
```
#SpawnItem BP_AK47
#SpawnVehicle BPC_Barba
```

## 模糊搜索说明

插件支持以下模糊搜索方式：

1. **按名称搜索**: `/scum 刀` 会匹配所有名称包含"刀"的物品
2. **按代码搜索**: `/scum BP_AK` 会匹配所有代码包含"BP_AK"的物品  
3. **按分类搜索**: `/scum 武器` 会匹配所有武器分类的物品
4. **按子分类搜索**: `/scum 摩托` 会匹配所有摩托车载具

## 数据来源

物品数据来源于 `SCUM全物品代码大全_豆包AI生成 (1).xlsx`，包含：
- **2247个物品**
- **多级分类**（大分类 → 小分类）
- **完整刷取指令**

## 文件结构

```
qqadmin_scum_search/
├── main.py                    # AstrBot插件主入口
├── metadata.yaml              # 插件元数据
├── _conf_schema.json          # 配置schema
├── __init__.py                # Python包结构
├── requirements.txt           # 依赖列表
├── README.md                  # 使用说明
└── SCUM全物品代码大全_豆包AI生成 (1).xlsx  # 物品数据库
```

## 常见问题

**Q: 插件没有显示在列表中？**
A: 确保 `metadata.yaml` 文件存在且格式正确。

**Q: 命令无响应？**
A: 检查 AstrBot 日志，确保插件已正确加载，Excel文件路径正确。

**Q: 查询结果太多？**
A: 在配置中减小 `max_results` 值。

**Q: 如何更新物品数据？**
A: 替换 `SCUM全物品代码大全_豆包AI生成 (1).xlsx` 文件并重启插件。

## 许可证

MIT License