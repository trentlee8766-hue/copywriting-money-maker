#!/usr/bin/env python3
"""
安全替换脚本
功能：
1. 按内容匹配替换文本，保持格式和样式
2. 支持批量替换（从JSON文件加载替换规则）
3. 支持AI特征词汇替换
4. 支持批注内容替换
"""

import os
import sys
import json
import argparse
import re
from typing import List, Dict, Tuple, Optional

# 添加模块路径
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    import win32com.client
    WIN32COM_AVAILABLE = True
except ImportError:
    WIN32COM_AVAILABLE = False
    print("警告：win32com不可用，无法执行Word替换操作")


class SafeReplacer:
    """安全替换器"""
    
    def __init__(self):
        self.stats = {
            "total_replacements": 0,
            "successful": 0,
            "failed": 0,
            "style_recovered": 0
        }
    
    def load_replacements(self, rules_path: str) -> List[Dict[str, str]]:
        """
        加载替换规则
        
        Args:
            rules_path: 规则文件路径（JSON）
            
        Returns:
            替换规则列表
        """
        try:
            with open(rules_path, 'r', encoding='utf-8') as f:
                rules_data = json.load(f)

            # 规则文件可能是分层结构（high_risk/medium_risk -> 字典）
            # 也可能是扁平字典形式（关键词->替换词）
            # 也可能是列表形式（old->new）
            replacements = []

            if isinstance(rules_data, dict):
                # 检查是否是分层结构（high_risk/medium_risk）
                first_key = next(iter(rules_data.keys()), None)
                if first_key in ('high_risk', 'medium_risk', 'low_risk'):
                    # 分层结构，展平
                    for risk_level, rules in rules_data.items():
                        if isinstance(rules, dict):
                            for old_text, new_text in rules.items():
                                replacements.append({
                                    "old": old_text,
                                    "new": new_text,
                                    "risk": risk_level
                                })
                else:
                    # 扁平结构
                    for old_text, new_text in rules_data.items():
                        replacements.append({
                            "old": old_text,
                            "new": new_text
                        })
            elif isinstance(rules_data, list):
                replacements = rules_data

            print(f"加载了 {len(replacements)} 条替换规则")
            return replacements
            
        except Exception as e:
            print(f"加载替换规则失败: {e}")
            return []
    
    def safe_replace_keep_style(self, doc_path: str, replacements: List[Dict[str, str]], 
                               output_path: str) -> Dict[str, int]:
        """
        安全段落替换：按内容匹配，保留格式和样式
        
        Args:
            doc_path: 原始文档路径
            replacements: 替换规则列表
            output_path: 输出文档路径
            
        Returns:
            替换统计信息
        """
        if not WIN32COM_AVAILABLE:
            print("错误：win32com不可用，无法执行替换")
            return self.stats
        
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(doc_path))
            
            for i, rep in enumerate(replacements):
                old_text = rep.get("old", "").strip()
                new_text = rep.get("new", "")
                
                if not old_text:
                    continue
                
                matched = False
                
                # 遍历所有段落
                for j in range(1, doc.Paragraphs.Count + 1):
                    para = doc.Paragraphs(j)
                    para_text = para.Range.Text.strip()
                    
                    # 模糊匹配：原文片段出现在段落中
                    if old_text in para_text:
                        # 保存原样式
                        try:
                            original_style = para.Style
                        except:
                            original_style = None
                        
                        # 关键：去掉段落标记再替换
                        range_obj = para.Range
                        try:
                            # 尝试去掉段落标记
                            if range_obj.Text.endswith('\r'):
                                range_obj.End = range_obj.End - 1
                            
                            range_obj.Text = new_text
                            
                            # 检查并恢复样式
                            if original_style and para.Style != original_style:
                                try:
                                    para.Style = original_style
                                    self.stats["style_recovered"] += 1
                                    print(f"⚠️ 样式恢复 [{i+1}]: P{j} 样式恢复为 {original_style.NameLocal}")
                                except:
                                    pass
                            
                            self.stats["successful"] += 1
                            matched = True
                            print(f"✅ 替换成功 [{i+1}]: P{j} | {old_text[:30]}...")
                            break  # 找到第一个匹配就跳出
                            
                        except Exception as e:
                            print(f"❌ 替换失败 [{i+1}]: P{j} 错误: {e}")
                            self.stats["failed"] += 1
                            break
                
                if not matched:
                    print(f"❌ 未匹配 [{i+1}]: {old_text[:50]}...")
                    self.stats["failed"] += 1
                
                self.stats["total_replacements"] += 1
            
            # 保存文档
            doc.SaveAs(os.path.abspath(output_path))
            doc.Close()
            word.Quit()
            
            print(f"\n替换统计:")
            print(f"  总计尝试: {self.stats['total_replacements']}")
            print(f"  成功: {self.stats['successful']}")
            print(f"  失败: {self.stats['failed']}")
            print(f"  样式恢复: {self.stats['style_recovered']}")
            
            return self.stats
            
        except Exception as e:
            print(f"文档处理失败: {e}")
            return self.stats
    
    def replace_ai_keywords(self, doc_path: str, keywords_path: str, 
                           rules_path: str, output_path: str) -> Dict[str, int]:
        """
        替换AI特征词汇
        
        Args:
            doc_path: 文档路径
            keywords_path: AI特征词库路径
            rules_path: 替换规则路径
            output_path: 输出路径
            
        Returns:
            替换统计
        """
        # 加载AI关键词
        try:
            with open(keywords_path, 'r', encoding='utf-8') as f:
                keywords_data = json.load(f)
            
            # 提取高风险和中风险词汇
            high_risk = keywords_data.get("high_risk", [])
            medium_risk = keywords_data.get("medium_risk", [])
            
            ai_keywords = high_risk + medium_risk
            
        except Exception as e:
            print(f"加载AI关键词失败: {e}")
            ai_keywords = []
        
        # 加载替换规则
        replacements = self.load_replacements(rules_path)
        
        # 筛选AI相关的替换规则
        ai_replacements = []
        for rep in replacements:
            old_text = rep.get("old", "")
            # 如果旧文本在AI关键词中，则加入
            for keyword in ai_keywords:
                if keyword in old_text:
                    ai_replacements.append(rep)
                    break
        
        print(f"找到 {len(ai_replacements)} 条AI相关替换规则")
        
        # 执行替换
        if ai_replacements:
            return self.safe_replace_keep_style(doc_path, ai_replacements, output_path)
        else:
            print("没有找到AI相关替换规则，跳过AI替换")
            return {"total": 0, "successful": 0, "failed": 0}
    
    def replace_comments(self, doc_path: str, output_path: str) -> Dict[str, int]:
        """
        处理批注（根据批注内容进行替换）
        
        Args:
            doc_path: 文档路径
            output_path: 输出路径
            
        Returns:
            处理统计
        """
        if not WIN32COM_AVAILABLE:
            print("错误：win32com不可用，无法处理批注")
            return {"total": 0, "successful": 0, "failed": 0}
        
        try:
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(os.path.abspath(doc_path))
            
            # 收集批注信息
            comments = []
            if hasattr(doc, 'Comments'):
                for i, comment in enumerate(doc.Comments):
                    try:
                        comment_text = comment.Range.Text if hasattr(comment.Range, 'Text') else ""
                        scope_text = comment.Scope.Text if hasattr(comment, 'Scope') else ""
                        author = comment.Author if hasattr(comment, 'Author') else "Unknown"
                        
                        comments.append({
                            "index": i + 1,
                            "author": author,
                            "comment": comment_text,
                            "scope": scope_text
                        })
                    except:
                        comments.append({"index": i + 1, "error": "无法读取"})
            
            print(f"找到 {len(comments)} 条批注")
            
            # 分析批注内容，生成替换规则
            replacements = []
            for comment in comments:
                comment_text = comment.get("comment", "")
                scope_text = comment.get("scope", "")
                
                # 简单的批注类型判断
                if "需要修改" in comment_text or "改为" in comment_text:
                    # 提取修改建议
                    # 这里需要更复杂的自然语言处理
                    # 暂时跳过，只记录
                    print(f"批注[{comment['index']}]: {comment_text[:50]}...")
            
            # 保存文档（可以选择删除批注）
            # 暂时不删除批注，只保存副本
            doc.SaveAs(os.path.abspath(output_path))
            doc.Close()
            word.Quit()
            
            return {"total": len(comments), "processed": 0, "skipped": len(comments)}
            
        except Exception as e:
            print(f"批注处理失败: {e}")
            return {"total": 0, "successful": 0, "failed": 0}
    
    def batch_replace(self, doc_path: str, replacements: List[Dict[str, str]], 
                     output_path: str, batch_size: int = 10) -> Dict[str, int]:
        """
        批量替换（分批处理，避免内存问题）
        
        Args:
            doc_path: 文档路径
            replacements: 替换规则列表
            output_path: 输出路径
            batch_size: 每批大小
            
        Returns:
            替换统计
        """
        total_batches = (len(replacements) + batch_size - 1) // batch_size
        all_stats = {"total": 0, "successful": 0, "failed": 0, "style_recovered": 0}
        
        for batch_num in range(total_batches):
            start = batch_num * batch_size
            end = min(start + batch_size, len(replacements))
            batch = replacements[start:end]
            
            print(f"\n处理批次 {batch_num + 1}/{total_batches} ({len(batch)}条规则)")
            
            # 每批使用临时文件
            if batch_num == 0:
                input_file = doc_path
            else:
                input_file = output_path.replace('.docx', f'_batch{batch_num}.docx')
            
            output_file = output_path.replace('.docx', f'_batch{batch_num+1}.docx')
            
            batch_stats = self.safe_replace_keep_style(input_file, batch, output_file)
            
            # 累加统计
            for key in all_stats:
                all_stats[key] += batch_stats.get(key, 0)
            
            # 清理临时文件（除了最后一个）
            if batch_num > 0 and os.path.exists(input_file):
                try:
                    os.remove(input_file)
                except:
                    pass
        
        # 重命名最终文件
        if total_batches > 1:
            final_temp = output_path.replace('.docx', f'_batch{total_batches}.docx')
            if os.path.exists(final_temp):
                os.rename(final_temp, output_path)
        
        return all_stats


def main():
    """命令行入口"""
    parser = argparse.ArgumentParser(description="安全替换脚本")
    
    # 替换模式
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--replace", action="store_true", help="执行普通替换")
    group.add_argument("--ai-replace", action="store_true", help="替换AI特征词汇")
    group.add_argument("--comment-replace", action="store_true", help="处理批注")
    
    # 通用参数
    parser.add_argument("--input", "-i", type=str, required=True, help="输入文档路径")
    parser.add_argument("--output", "-o", type=str, required=True, help="输出文档路径")
    parser.add_argument("--rules", "-r", type=str, help="替换规则文件路径（JSON）")
    parser.add_argument("--keywords", "-k", type=str, help="AI特征词库路径（JSON）")
    parser.add_argument("--batch-size", "-b", type=int, default=10, help="批量处理大小")
    
    args = parser.parse_args()
    
    replacer = SafeReplacer()
    
    if args.ai_replace:
        # AI替换模式
        if not args.keywords:
            print("错误：AI替换模式需要--keywords参数")
            return
        
        if not args.rules:
            print("错误：AI替换模式需要--rules参数")
            return
        
        stats = replacer.replace_ai_keywords(
            args.input, args.keywords, args.rules, args.output
        )
        
    elif args.comment_replace:
        # 批注处理模式
        stats = replacer.replace_comments(args.input, args.output)
        
    else:
        # 普通替换模式
        if not args.rules:
            print("错误：普通替换模式需要--rules参数")
            return
        
        replacements = replacer.load_replacements(args.rules)
        if not replacements:
            print("错误：未加载到替换规则")
            return
        
        if len(replacements) > args.batch_size:
            stats = replacer.batch_replace(
                args.input, replacements, args.output, args.batch_size
            )
        else:
            stats = replacer.safe_replace_keep_style(
                args.input, replacements, args.output
            )
    
    # 输出统计
    print("\n" + "=" * 50)
    print("替换完成!")
    for key, value in stats.items():
        print(f"{key}: {value}")
    
    return stats


if __name__ == "__main__":
    main()