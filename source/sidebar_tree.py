# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import os
from unohelper import systemPathToFileUrl
from com.sun.star.awt.tree import XMutableTreeDataModel

class SidebarTree():

    def __init__(self, ctx):
        self.ctx = ctx
        self.svg_url = None
        self.category_name = None
        self.category_dir_path = None
        self.favorites_dir_path = None

    def set_favorites_dir_path(self, favorites_dir_path):
        self.favorites_dir_path = favorites_dir_path

    def set_category_name(self, name):
        self.category_name = name
        self.create_category_dir_path(name)

    def create_category_dir_path(self, category_name):
        self.category_dir_path = os.path.join(self.favorites_dir_path, category_name)
        os.makedirs(self.category_dir_path, exist_ok=True)

    def set_svg_data(self, svg_data):
        self.svg_url = self.generate_svg_file_url(svg_data)

    def generate_svg_file_url(self, svg_data):
        counter = 1
        name = "Symbol"
        while True:
            self.name = f"{name} {counter}"
            file_name = self.name + ".svg"
            symbol_svg_path = os.path.join(self.category_dir_path, file_name)
            if not os.path.exists(symbol_svg_path):
                break
            counter += 1

        with open(symbol_svg_path, 'w', encoding='utf-8') as preview_file:
            preview_file.write(svg_data)

        return systemPathToFileUrl(symbol_svg_path)

    def refresh_tree(self, root_node, tree_data_model):
        try:
            existing_category_node = None
            for i in range(root_node.getChildCount()):
                child = root_node.getChildAt(i)
                if child.getDisplayValue() == self.category_name:
                    existing_category_node = child
                    break

            if existing_category_node is None:
                existing_category_node = tree_data_model.createNode(self.category_name, True)
                root_node.appendChild(existing_category_node)

            symbol_child = tree_data_model.createNode(self.name, False)
            symbol_child.setNodeGraphicURL(self.svg_url)

            existing_category_node.appendChild(symbol_child)
        except Exception as e:
            print("refresh_tree error:", e)
        