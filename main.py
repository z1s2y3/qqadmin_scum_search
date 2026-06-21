"""
SCUM物品代码查询插件 - AstrBot版本
支持模糊搜索和官方中文名称
数据来源: SCUM全物品代码大全_豆包AI生成.xlsx
"""

import asyncio
import openpyxl
import os
from typing import List, Dict, Optional
from astrbot.api.event import filter, AstrMessageEvent
from astrbot.api.star import Context, Star, register
from astrbot.api import logger

SCUM_ITEMS: List[Dict[str, str]] = []
MAIN_CATEGORIES: List[str] = []
SUB_CATEGORIES: Dict[str, List[str]] = {}

OFFICIAL_NAMES = {
    "简易手推车": "简易手推车",
    "金属手推车": "金属手推车",
    "巴巴摩托": "巴巴摩托",
    "AK47突击步枪": "AK-47",
    "AKM突击步枪": "AKM",
    "M4A1突击步枪": "M4A1",
    "M16A4突击步枪": "M16A4",
    "SCAR-H突击步枪": "SCAR-H",
    "格洛克17": "Glock 17",
    "格洛克21": "Glock 21",
    "沙漠之鹰": "Desert Eagle",
    "M9手枪": "M9",
    "1911手枪": "M1911",
    "左轮手枪": "Revolver",
    "莫辛纳甘": "Mosin-Nagant",
    "SKS步枪": "SKS",
    "SV98狙击枪": "SV-98",
    "VSS狙击枪": "VSS",
    "AWP狙击枪": "AWP",
    "M24狙击枪": "M24",
    "MP5冲锋枪": "MP5",
    "UMP45冲锋枪": "UMP45",
    "P90冲锋枪": "P90",
    "MAC10冲锋枪": "MAC-10",
    "Vector冲锋枪": "Vector",
    "PPSh41冲锋枪": "PPSh-41",
    "M249轻机枪": "M249",
    "PKM机枪": "PKM",
    "RPK机枪": "RPK",
    "M870霰弹枪": "M870",
    "SPAS12霰弹枪": "SPAS-12",
    "复合弓": "Compound Bow",
    "十字弩": "Crossbow",
    "砍刀": "Machete",
    "战术小刀": "Tactical Knife",
    "撬棍": "Crowbar",
    "棒球棒": "Baseball Bat",
    "工兵铲": "Shovel",
}

def load_items():
    excel_path = os.path.join(os.path.dirname(__file__), "SCUM全物品代码大全_豆包AI生成 (1).xlsx")
    if not os.path.exists(excel_path):
        logger.error(f"Excel文件不存在: {excel_path}")
        return
    
    try:
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active
        
        for row in range(2, sheet.max_row + 1):
            try:
                main_cat = sheet.cell(row=row, column=1).value or ""
                sub_cat = sheet.cell(row=row, column=2).value or ""
                code = sheet.cell(row=row, column=3).value or ""
                name = sheet.cell(row=row, column=4).value or ""
                
                if not code:
                    continue
                
                SCUM_ITEMS.append({
                    "main_category": main_cat.strip() if main_cat else "",
                    "sub_category": sub_cat.strip() if sub_cat else "",
                    "code": code.strip(),
                    "name": name.strip() if name else ""
                })
                
                if main_cat and main_cat.strip() not in MAIN_CATEGORIES:
                    MAIN_CATEGORIES.append(main_cat.strip())
                
                if main_cat and sub_cat:
                    main_key = main_cat.strip()
                    sub_key = sub_cat.strip()
                    if main_key not in SUB_CATEGORIES:
                        SUB_CATEGORIES[main_key] = []
                    if sub_key not in SUB_CATEGORIES[main_key]:
                        SUB_CATEGORIES[main_key].append(sub_key)
            except Exception as e:
                logger.error(f"读取第{row}行失败: {e}")
        
        logger.info(f"已加载 {len(SCUM_ITEMS)} 个物品")
    except Exception as e:
        logger.error(f"加载Excel失败: {e}")

def fuzzy_search(keyword: str) -> List[Dict[str, str]]:
    results = []
    keyword = keyword.lower().strip()
    
    for item in SCUM_ITEMS:
        name = item["name"].lower() if item["name"] else ""
        code = item["code"].lower() if item["code"] else ""
        sub_cat = item["sub_category"].lower() if item["sub_category"] else ""
        
        if keyword in name or keyword in code or keyword in sub_cat:
            results.append(item)
    
    return sorted(results, key=lambda x: (len(x["name"]), x["name"]))

def search_by_sub_category(sub_category: str) -> List[Dict[str, str]]:
    results = []
    sub_category = sub_category.lower().strip()
    
    for item in SCUM_ITEMS:
        item_sub_cat = item["sub_category"].lower() if item["sub_category"] else ""
        if sub_category == item_sub_cat:
            results.append(item)
    
    return sorted(results, key=lambda x: x["name"] if x["name"] else x["code"])

def get_category_items(main_category: str) -> List[Dict[str, str]]:
    results = []
    main_category = main_category.lower().strip()
    
    for item in SCUM_ITEMS:
        item_main_cat = item["main_category"].lower() if item["main_category"] else ""
        if main_category == item_main_cat:
            results.append(item)
    
    return sorted(results, key=lambda x: (x["sub_category"], x["name"]))

def get_official_name(name: str) -> str:
    for key, value in OFFICIAL_NAMES.items():
        if key in name or name in key:
            return value
    return name

@register
class ScumItemSearch(Star):
    name: str = "SCUM物品查询"
    description: str = "查询SCUM游戏物品代码"
    version: str = "1.0.0"
    author: str = "AstrBot"
    
    async def initialize(self, ctx: Context):
        load_items()
    
    @filter.command("scum")
    async def scum_search(self, event: AstrMessageEvent):
        args = event.message_str.strip().split()
        if len(args) < 2:
            await self.scum_help(event)
            return
        
        keyword = ' '.join(args[1:])
        results = fuzzy_search(keyword)
        
        if not results:
            response = f"❌ 未找到「{keyword}」相关物品"
            yield event.plain_result(response)
            return
        
        response = f"🔍 查询结果 ({len(results)}个):\n\n"
        for i, item in enumerate(results[:10], 1):
            name = item["name"]
            official_name = get_official_name(name)
            if official_name != name and official_name:
                display_name = f"{name} ({official_name})"
            else:
                display_name = name if name else item["code"]
            
            response += f"📦 {display_name}\n"
            response += f"   ├─ 分类: {item['main_category']} > {item['sub_category']}\n"
            response += f"   └─ 代码: {item['code']}\n"
            response += f"   🎮 #SpawnItem {item['code']}\n\n"
        
        if len(results) > 10:
            response += f"💡 还有 {len(results) - 10} 个结果，建议使用更精确的关键词"
        
        yield event.plain_result(response)
    
    @filter.command("scum分类")
    async def scum_categories(self, event: AstrMessageEvent):
        response = "📦 SCUM物品分类\n\n"
        
        for main_cat in MAIN_CATEGORIES:
            response += f"• {main_cat}\n"
            if main_cat in SUB_CATEGORIES:
                for sub_cat in SUB_CATEGORIES[main_cat]:
                    response += f"  └─ {sub_cat}\n"
        
        response += f"\n💡 使用: /scum子分类 <子分类名称>\n示例: /scum子分类 武器"
        yield event.plain_result(response)
    
    @filter.command("scum子分类")
    async def scum_sub_category(self, event: AstrMessageEvent):
        args = event.message_str.strip().split()
        if len(args) < 2:
            response = "❌ 请指定子分类名称\n\n💡 使用: /scum子分类 <名称>\n示例: /scum子分类 武器"
            yield event.plain_result(response)
            return
        
        sub_category = ' '.join(args[1:])
        items = search_by_sub_category(sub_category)
        
        if not items:
            response = f"❌ 未找到「{sub_category}」子分类"
            yield event.plain_result(response)
            return
        
        response = f"📦 {sub_category} ({len(items)}个):\n\n"
        for i, item in enumerate(items[:10], 1):
            name = item["name"]
            official_name = get_official_name(name)
            if official_name != name and official_name:
                display_name = f"{name} ({official_name})"
            else:
                display_name = name if name else item["code"]
            
            response += f"{i}. {display_name} - {item['code']}\n"
        
        if len(items) > 10:
            response += f"\n💡 还有 {len(items) - 10} 个物品，使用 /scum <关键词> 搜索"
        
        yield event.plain_result(response)
    
    @filter.command("scum帮助")
    async def scum_help(self, event: AstrMessageEvent):
        response = """🎮 SCUM物品查询
        
📌 命令:
├─ /scum <关键词> - 模糊搜索物品
├─ /scum分类 - 查看所有分类
├─ /scum子分类 <名称> - 按子分类查询
└─ /scum帮助 - 显示此帮助

🎯 使用示例:
/scum ak47         # 搜索AK-47
/scum 凤凰针       # 搜索凤凰针
/scum 巴雷特礼盒   # 搜索巴雷特礼盒
/scum 刀           # 模糊搜索所有刀类
/scum子分类 武器    # 查看武器分类

💡 模糊搜索支持:
- 物品名称（如"凤凰针"、"AK47"）
- 物品代码（如"BP_AK47"）
- 分类名称（如"武器"、"医疗"）

🎮 游戏内使用:
在游戏控制台输入 #SpawnItem <代码>
例如: #SpawnItem BP_AK47

📦 数据统计:
- 总物品数: 2247个
- 大分类: 物品、载具
- 子分类: 16个"""
        
        yield event.plain_result(response)
    
    async def terminate(self):
        logger.info("SCUM物品查询插件已卸载")