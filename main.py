"""
SCUM物品代码查询插件 - AstrBot版本
支持模糊搜索和官方中文名称
数据来源: SCUM全物品代码大全_豆包AI生成.xlsx
"""

import openpyxl
import os
from typing import List, Dict
from astrbot.api.event import filter, AstrMessageEvent
from astrbot.api.star import Star, register
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
    "Weapon_AK47": "AK-47突击步枪",
    "Weapon_AKM": "AKM突击步枪",
    "Weapon_M4A1": "M4A1突击步枪",
    "Weapon_M16A4": "M16A4突击步枪",
    "Weapon_SCAR_H": "SCAR-H突击步枪",
    "Weapon_Glock17": "格洛克17手枪",
    "Weapon_Glock21": "格洛克21手枪",
    "Weapon_DesertEagle": "沙漠之鹰手枪",
    "Weapon_M9": "M9手枪",
    "Weapon_1911": "M1911手枪",
    "Weapon_Revolver": "左轮手枪",
    "Weapon_MosinNagant": "莫辛纳甘步枪",
    "Weapon_SKS": "SKS步枪",
    "Weapon_SV98": "SV-98狙击枪",
    "Weapon_VSS": "VSS狙击枪",
    "Weapon_AWP": "AWP狙击枪",
    "Weapon_M24": "M24狙击枪",
    "Weapon_MP5": "MP5冲锋枪",
    "Weapon_UMP45": "UMP45冲锋枪",
    "Weapon_P90": "P90冲锋枪",
    "Weapon_MAC10": "MAC-10冲锋枪",
    "Weapon_Vector": "Vector冲锋枪",
    "Weapon_PPSh41": "PPSh-41冲锋枪",
    "Weapon_M249": "M249轻机枪",
    "Weapon_PKM": "PKM机枪",
    "Weapon_RPK": "RPK机枪",
    "Weapon_M870": "M870霰弹枪",
    "Weapon_SPAS12": "SPAS-12霰弹枪",
    "Weapon_Machete": "砍刀",
    "Weapon_Knife": "战术小刀",
    "Weapon_Crowbar": "撬棍",
    "Weapon_BaseballBat": "棒球棒",
    "Weapon_Shovel": "工兵铲",
    "Magazine_AK47": "AK47弹夹",
    "Magazine_M4A1": "M4A1弹夹",
    "Magazine_M16A4": "M16A4弹夹",
    "ScopeRail_AK47": "AK47导轨",
    "Weapon_Parts_AK47": "AK47零件",
    "Weapon_AK47_Engraved": "刻字AK-47",
}

def load_items(excel_path=None):
    if excel_path is None:
        excel_path = os.path.join(os.path.dirname(__file__), "SCUM_CN.xlsx")
    else:
        excel_path = os.path.join(os.path.dirname(__file__), excel_path)
    
    print(f"[SCUM插件] 尝试加载Excel: {excel_path}")
    
    if not os.path.exists(excel_path):
        logger.error(f"Excel文件不存在: {excel_path}")
        print(f"[SCUM插件] 错误：Excel文件不存在!")
        return
    
    try:
        print(f"[SCUM插件] 正在打开Excel文件...")
        wb = openpyxl.load_workbook(excel_path)
        sheet = wb.active
        print(f"[SCUM插件] Excel文件已打开，共 {sheet.max_row} 行")
        
        headers = [sheet.cell(row=1, column=i).value for i in range(1, 4)]
        logger.info(f"Excel表头: {headers}")
        print(f"[SCUM插件] 表头: {headers}")
        
        for row in range(2, sheet.max_row + 1):
            try:
                name = sheet.cell(row=row, column=1).value or ""
                code = sheet.cell(row=row, column=2).value or ""
                spawn_cmd = sheet.cell(row=row, column=3).value or ""
                
                if not code:
                    continue
                
                main_cat = "物品"
                sub_cat = ""
                
                if "Vehicle" in code or "Car" in code or "Truck" in code or "Boat" in code or "Bike" in code:
                    main_cat = "载具"
                    sub_cat = "载具"
                elif "Zombie" in code:
                    main_cat = "丧尸"
                    sub_cat = "丧尸"
                elif "Weapon_" in code:
                    main_cat = "物品"
                    sub_cat = "武器"
                elif "Magazine_" in code:
                    main_cat = "物品"
                    sub_cat = "弹药"
                elif "Ammo" in code or "Cal_" in code or "Gauge" in code:
                    main_cat = "物品"
                    sub_cat = "弹药"
                elif "Armor" in code or "Cloth" in code or "Shirt" in code or "Pants" in code or "Boots" in code:
                    main_cat = "物品"
                    sub_cat = "护甲"
                elif "Food" in code or "Water" in code or "Drink" in code:
                    main_cat = "物品"
                    sub_cat = "食物"
                elif "Med" in code or "Heal" in code or "Bandage" in code or "Syringe" in code:
                    main_cat = "物品"
                    sub_cat = "医疗"
                elif "Tool" in code or "Part" in code or "Component" in code:
                    main_cat = "物品"
                    sub_cat = "工具"
                elif "Build" in code or "Structure" in code or "Fence" in code or "Wall" in code:
                    main_cat = "物品"
                    sub_cat = "建家"
                else:
                    main_cat = "物品"
                    sub_cat = "其他"
                
                SCUM_ITEMS.append({
                    "main_category": main_cat.strip(),
                    "sub_category": sub_cat.strip(),
                    "code": code.strip(),
                    "name": name.strip() if name else "",
                    "spawn_cmd": spawn_cmd.strip() if spawn_cmd else ""
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
        print(f"[SCUM插件] 已加载 {len(SCUM_ITEMS)} 个物品")
        if SCUM_ITEMS:
            print(f"[SCUM插件] 示例数据: {SCUM_ITEMS[0]}")
    except Exception as e:
        logger.error(f"加载Excel失败: {e}")
        print(f"[SCUM插件] 加载Excel失败: {e}")
        import traceback
        traceback.print_exc()

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
    if not name:
        return ""
    
    if name in OFFICIAL_NAMES:
        return OFFICIAL_NAMES[name]
    
    for key, value in OFFICIAL_NAMES.items():
        if key in name or name in key:
            return value
    return name

# 模块加载时立即初始化数据
print("[SCUM插件] 正在初始化...")
load_items()
print(f"[SCUM插件] 初始化完成，共 {len(SCUM_ITEMS)} 个物品")

@register("SCUM物品查询", "AstrBot", "查询SCUM游戏物品代码", "1.0.0")
class ScumItemSearch(Star):
    
    def __init__(self, context, config=None):
        super().__init__(context)
        self.config = config or {}
    
    async def on_load(self):
        try:
            logger.info("开始加载SCUM物品数据...")
            excel_path = self.config.get('excel_path', None)
            print(f"[SCUM插件] 从配置读取excel_path: {excel_path}")
            load_items(excel_path)
            logger.info(f"SCUM物品加载完成，共 {len(SCUM_ITEMS)} 个物品")
        except Exception as e:
            logger.error(f"加载SCUM物品数据失败: {e}")
            import traceback
            logger.error(traceback.format_exc())
    
    def get_spawn_command(self, item):
        # 优先使用Excel中的完整刷取指令
        if item.get("spawn_cmd"):
            return item["spawn_cmd"]
        
        # 如果没有则自动生成
        main_cat = item["main_category"]
        code = item["code"]
        
        if main_cat == "载具":
            return f"#SpawnVehicle {code} 1"
        elif main_cat == "丧尸":
            return f"#SpawnZombie {code} 1"
        elif "Magazine" in code or "Ammo" in code:
            return f"#SpawnItem {code} 1 StackCount 100"
        else:
            return f"#SpawnItem {code} 1"
    
    def get_quantity_word(self, item):
        main_cat = item["main_category"]
        sub_cat = item["sub_category"]
        
        if main_cat == "载具":
            return "辆"
        elif sub_cat in ["武器", "单手近战", "双手近战"]:
            return "把"
        elif sub_cat in ["护甲", "服装", "饰品"]:
            return "件"
        elif sub_cat in ["弹药"]:
            return "发"
        elif sub_cat in ["食物", "饮料", "药品"]:
            return "个"
        else:
            return "个"
    
    @filter.command("scum")
    async def scum_search(self, event: AstrMessageEvent):
        args = event.message_str.strip().split()
        if len(args) < 2:
            await self.scum_help(event)
            return
        
        keyword = args[1]
        page = 1
        
        if len(args) >= 3 and args[2].isdigit():
            page = int(args[2])
        
        results = fuzzy_search(keyword)
        
        if not results:
            response = f"❌ 未找到「{keyword}」相关物品"
            yield event.plain_result(response)
            return
        
        max_results = self.config.get('max_results', 10)
        print(f"[SCUM插件] 当前max_results配置: {max_results}")
        
        total_pages = (len(results) + max_results - 1) // max_results
        start_idx = (page - 1) * max_results
        end_idx = start_idx + max_results
        
        page_results = results[start_idx:end_idx]
        
        if page > total_pages:
            response = f"❌ 没有第 {page} 页，共 {total_pages} 页"
            yield event.plain_result(response)
            return
        
        response = "══════════════════════════════════════════\n"
        
        for item in page_results:
            name = item["name"]
            official_name = get_official_name(name)
            code = item["code"]
            main_cat = item["main_category"]
            
            if official_name:
                display_name = official_name
            else:
                display_name = name if name else code
            
            cat_icon = "📦"
            if main_cat == "载具":
                cat_icon = "🚗"
            elif main_cat == "丧尸":
                cat_icon = "🧟"
            elif "武器" in item["sub_category"] or "近战" in item["sub_category"]:
                cat_icon = "🔫"
            elif "弹药" in item["sub_category"]:
                cat_icon = "💣"
            
            response += f"{cat_icon} 【{display_name}】\n"
            response += f"   └─ {main_cat} > {item['sub_category']}\n"
            response += f"   ├─ 物品代码: {code}\n"
            response += f"   └─ 游戏代码: 🎮 #SpawnItem {code} 1\n"
            
            if "Magazine" in code or "Ammo" in code:
                response += f"   └─ 满弹夹: 🎮 #SpawnItem {code} 1 StackCount 100\n"
            
            response += "──────────────────────────────────────────\n"
        
        if total_pages > 1:
            response += f"📄 第 {page}/{total_pages} 页"
            if page < total_pages:
                response += f" | 输入 /scum {keyword} {page + 1} 查看下一页"
            else:
                response += " | 已是最后一页"
        
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