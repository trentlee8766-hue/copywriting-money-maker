#!/usr/bin/env python3
"""
AI特征检测模块
功能：
1. 加载AI特征词库（config/ai_keywords.json）
2. 检测文本中的AI特征词汇密度
3. 预判AI率
4. 生成替换建议
"""

import json
import re
import sys
import os
from typing import Dict, List, Tuple

# 添加当前目录到路径，以便导入其他模块
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class AIDetector:
    """AI特征检测器"""
    
    def __init__(self, keywords_path: str = None):
        """
        初始化检测器
        
        Args:
            keywords_path: AI特征词库路径，默认为config/ai_keywords.json
        """
        if keywords_path is None:
            # 默认路径
            base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            keywords_path = os.path.join(base_dir, "config", "ai_keywords.json")
        
        self.keywords_path = keywords_path
        self.keywords = self._load_keywords()
        
    def _load_keywords(self) -> Dict[str, List[str]]:
        """加载AI特征词库"""
        try:
            with open(self.keywords_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except FileNotFoundError:
            print(f"警告：AI特征词库文件不存在 {self.keywords_path}")
            return {
                "high_risk": ["旨在", "深入研究", "具有重要意义"],
                "medium_risk": ["具体而言", "由此可见"],
                "low_risk": ["然而", "但是"]
            }
    
    def detect_paragraph(self, text: str) -> Dict[str, int]:
        """
        检测单个段落中的AI特征词汇
        
        Returns:
            字典，包含各风险等级的词汇计数
        """
        result = {
            "high_risk": 0,
            "medium_risk": 0,
            "low_risk": 0,
            "total_words": len(text)
        }
        
        # 遍历各风险等级的词汇
        for risk_level, words in self.keywords.items():
            for word in words:
                # 简单计数，考虑词汇可能包含省略号等情况
                if "..." in word:
                    # 处理模式如"一是...二是...三是"
                    pattern = word.replace("...", ".*?")
                    count = len(re.findall(pattern, text))
                else:
                    count = text.count(word)
                
                result[risk_level] += count
        
        return result
    
    def detect_document(self, paragraphs: List[str]) -> List[Dict[str, int]]:
        """
        检测整个文档
        
        Args:
            paragraphs: 段落列表
            
        Returns:
            每个段落的检测结果列表
        """
        results = []
        for i, para in enumerate(paragraphs):
            if para.strip():  # 跳过空段落
                para_result = self.detect_paragraph(para)
                para_result["paragraph_index"] = i + 1
                para_result["text_preview"] = para[:50] + "..." if len(para) > 50 else para
                results.append(para_result)
        
        return results
    
    def predict_ai_rate(self, detection_results: List[Dict[str, int]]) -> float:
        """
        预判AI率
        
        Args:
            detection_results: 检测结果列表
            
        Returns:
            预判的AI率（0-100）
        """
        if not detection_results:
            return 0.0
        
        total_high = sum(r["high_risk"] for r in detection_results)
        total_medium = sum(r["medium_risk"] for r in detection_results)
        total_low = sum(r["low_risk"] for r in detection_results)
        total_paragraphs = len(detection_results)
        
        # 简单的加权算法
        # 高风险词汇权重高，中风险中等，低风险低
        ai_score = total_high * 3 + total_medium * 2 + total_low * 1
        
        # 归一化到0-100，基于段落数调整
        normalized_score = min(100, ai_score * 10 / max(1, total_paragraphs))
        
        return round(normalized_score, 2)
    
    def generate_report(self, detection_results: List[Dict[str, int]], 
                       ai_rate: float) -> str:
        """
        生成AI率预判报告
        
        Args:
            detection_results: 检测结果
            ai_rate: 预判AI率
            
        Returns:
            报告文本
        """
        report_lines = []
        report_lines.append("【AI率预判报告】")
        report_lines.append("=" * 50)
        report_lines.append(f"预判AI率: {ai_rate}%")
        report_lines.append("")
        
        # 按段落统计
        high_risk_paras = []
        medium_risk_paras = []
        
        for result in detection_results:
            idx = result["paragraph_index"]
            high = result["high_risk"]
            medium = result["medium_risk"]
            
            if high >= 2:
                high_risk_paras.append((idx, high, result["text_preview"]))
            elif medium >= 2:
                medium_risk_paras.append((idx, medium, result["text_preview"]))
        
        # 高风险段落
        if high_risk_paras:
            report_lines.append("高风险段落（建议优先修改）:")
            for idx, count, preview in high_risk_paras[:10]:  # 只显示前10个
                report_lines.append(f"  P{idx}: 高风险词汇{count}个 | '{preview}'")
            if len(high_risk_paras) > 10:
                report_lines.append(f"  ... 还有{len(high_risk_paras)-10}个高风险段落")
            report_lines.append("")
        
        # 中风险段落
        if medium_risk_paras:
            report_lines.append("中风险段落（建议修改）:")
            for idx, count, preview in medium_risk_paras[:10]:
                report_lines.append(f"  P{idx}: 中风险词汇{count}个 | '{preview}'")
            if len(medium_risk_paras) > 10:
                report_lines.append(f"  ... 还有{len(medium_risk_paras)-10}个中风险段落")
            report_lines.append("")
        
        # 修改建议
        report_lines.append("修改建议:")
        report_lines.append("1. 优先处理高风险段落（每段≥2个高风险词汇）")
        report_lines.append("2. 使用替换规则库（config/replacement_rules.json）进行系统化替换")
        report_lines.append("3. 建议用户先用查重平台测真实AI率，根据反馈调整策略")
        report_lines.append("4. 降AI完成后务必进行格式验证")
        
        return "\n".join(report_lines)
    
    def get_replacement_suggestions(self, text: str) -> List[Tuple[str, str, str]]:
        """
        获取替换建议
        
        Args:
            text: 文本
            
        Returns:
            替换建议列表 (原词, 替换词, 风险等级)
        """
        suggestions = []
        
        # 加载替换规则
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        rules_path = os.path.join(base_dir, "config", "replacement_rules.json")
        
        try:
            with open(rules_path, 'r', encoding='utf-8') as f:
                rules = json.load(f)
        except FileNotFoundError:
            rules = {}
        
        # 检查每个高风险词汇
        for risk_level, words in self.keywords.items():
            if risk_level != "high_risk":
                continue
                
            for word in words:
                if word in text and word in rules:
                    suggestions.append((word, rules[word], risk_level))
        
        return suggestions


def main():
    """命令行入口"""
    import argparse
    
    parser = argparse.ArgumentParser(description="AI特征检测模块")
    parser.add_argument("--predict", action="store_true", help="预判AI率")
    parser.add_argument("--input", type=str, help="输入文件路径（支持 .docx 和 .txt）")
    parser.add_argument("--keywords", type=str, default=None, help="AI特征词库路径")
    parser.add_argument("--output", type=str, help="输出报告路径")
    
    args = parser.parse_args()
    
    # 创建检测器
    detector = AIDetector(args.keywords)
    
    # 读取输入文件（自动识别格式）
    if args.input and os.path.exists(args.input):
        if args.input.endswith('.docx'):
            # 使用 python-docx 读取 Word 文档
            from docx import Document
            doc = Document(args.input)
            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
        else:
            # 读取文本文件
            with open(args.input, 'r', encoding='utf-8') as f:
                content = f.read()
            paragraphs = [p.strip() for p in content.split('\n') if p.strip()]
    else:
        # 如果没有输入文件，使用示例文本
        print("警告：未提供输入文件，使用示例文本")
        paragraphs = [
            "本研究旨在深入研究小学教育中的游戏化教学，具有重要意义。",
            "具体而言，我们将从多维度分析游戏化教学的实践价值。",
            "然而，现有的研究还存在一些不足。"
        ]
    
    # 检测
    results = detector.detect_document(paragraphs)
    ai_rate = detector.predict_ai_rate(results)
    
    # 生成报告
    report = detector.generate_report(results, ai_rate)
    
    # 输出
    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(report)
        print(f"报告已保存到: {args.output}")
    else:
        print(report)
    
    # 返回AI率（供其他脚本使用）
    if args.predict:
        print(f"预判AI率: {ai_rate}%")
        return ai_rate


if __name__ == "__main__":
    main()