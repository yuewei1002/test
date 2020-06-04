# -*- coding: utf-8 -*-


"""封装常用的XML功能
支持文件和xml字符串的处理。
支持新增节点，tag, 属性，值。
支持修改节点的值，属性。
支持节点的删除，属性的删除。
支持生成xml文件。"""

import xml.etree.ElementTree as ET

class HandleXml:
    def read_xml_file(self, file_path):
        '''
        读取xml文件
        :param file_path:待读取的xml文件路径
        :return: element tree的根节点
        '''
        self.tree = ET.parse(file_path)
        return self.tree.getroot()

    def read_xml_string(self, string):
        '''
        导入xml字符串
        :param string: 待导入的xml字符串
        :return: 根节点
        '''
        self.root = ET.fromstring(string)
        return self.root

    def write_xml(self, out_path, root_node=None):
        '''
        写入xml文件
        :param out_path:待写入xml文件的路径
        :param root_node:指定写入文件的根节点
        :return: 无
        '''
        if root_node == None:
            self.tree.write(out_path, encoding="utf-8", xml_declaration=True)
        else:
            tree = ET.ElementTree(root_node)
            tree.write(out_path, encoding="utf-8", xml_declaration=True)

    def find_nodes_by_tag(self, root_node,tag_name):
        '''
        通过标签名称查找节点
        :param root_node:传入待查找节点的根节点
        :param tag_name: 二级标签名称或路径
        :return: 查找出的节点列表
        '''
        return root_node.iter(tag_name)

    @staticmethod
    def update_node_attrib(nodelist, attrib_map):
        '''
        通过节点修改节点的属性
        :param nodelist: 待修改的节点列表
        :param attrib_map: 待修改的节点属性，字典格式
        :return: 无
        '''
        for node in nodelist:
            for k, v in attrib_map.items():
                node.attrib[k] = v

    @staticmethod
    def update_node_text(nodelist, text):
        '''
        通过节点修改节点的内容
        :param nodelist: 待修改的节点列表
        :param text: 待修改的节点内容
        :return: 无
        '''
        for node in nodelist:
            node.text = text

    @staticmethod
    def add_node(tag, property_map=None, content=None):
        '''
        添加节点
        :param tag:节点标签名称
        :param property_map: 节点的属性，字典格式N
        :param content: 节点的内容
        :return: 节点元素
        '''
        if property_map != None and content != None:
            element = ET.Element(tag, attrib=property_map)
            element.text = content
        elif property_map != None and content == None:
            element = ET.Element(tag, attrib=property_map)
        elif property_map == None and content != None:
            element = ET.Element(tag)
            element.text = content
        else:
            element = ET.Element(tag)
        return element

    @staticmethod
    def add_child_node(node_father, tag, property_map=None, content=None):
        '''
        添加子节点
        '''
        if property_map != None and content != None:
            element = ET.SubElement(node_father, tag, attrib=property_map)
            element.text = content
        elif property_map != None and content == None:
            element = ET.SubElement(node_father, tag, attrib=property_map)
        elif property_map == None and content != None:
            element = ET.SubElement(node_father, tag)
            element.text = content
        else:
            element = ET.SubElement(node_father, tag)
        return element

    @staticmethod
    def del_node_by_attrib(nodelist, kv_map):
        '''
        通过属性及属性值定位一个节点，并删除
        :param nodelist:节点列表
        :param kv_map:属性值，字典格式
        :return:无
        '''
        for node in nodelist:
            for child in node:
                for k, v in kv_map.items():
                    if child.attrib.get(k) and child.attrib.get(k).find(v) >= 0:
                        node.remove(child)

    @staticmethod
    def del_node_attrib(nodelist, attrib_map):
        '''
        删除节点属性值
        :param nodelist: 节点列表
        :param attrib_map: 节点属性，字典格式
        :return: 无
        '''
        for node in nodelist:
            for key in attrib_map:
                if key in node.attrib:
                    del node.attrib[key]