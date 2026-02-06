# SPDX-FileCopyrightText: Collabora Productivity and contributors
#
# SPDX-License-Identifier: MPL-2.0
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

"""
Translator module for loading localized strings from property files
"""

import uno
from utils import get_package_location


class Translator:
    """Handles translation of strings from resource property files"""

    _instance = None
    _initialized = False

    def __new__(cls, x_context=None):
        """Singleton pattern to ensure only one Translator instance"""
        if cls._instance is None:
            cls._instance = super(Translator, cls).__new__(cls)
        return cls._instance

    def __init__(self, x_context):
        """Initialize translator with UNO context"""
        # Only initialize once (singleton pattern)
        if not Translator._initialized:
            self._x_context = x_context
            self._resource_cache = {}
            Translator._initialized = True

    def get_locale(self):
        """Get locale from configuration provider"""
        try:
            x_mcf = self._x_context.getServiceManager()
            o_configuration_provider = x_mcf.createInstanceWithContext(
                "com.sun.star.configuration.ConfigurationProvider", self._x_context
            )
            x_localizable = o_configuration_provider
            locale = x_localizable.getLocale()
            return locale
        except Exception as ex:
            print(f"Error getting locale: {ex}")
            return None

    def get_string_resource(self, dialog_name):
        """Get or create string resource for a dialog"""
        # Check cache first
        if dialog_name in self._resource_cache:
            return self._resource_cache[dialog_name]

        x_resources = None
        m_res_root_url = get_package_location(self._x_context) + "/dialog/"

        try:
            args = (
                m_res_root_url,
                True,
                self.get_locale(),
                dialog_name,
                "",
                uno.Any("com.sun.star.task.XInteractionHandler", None),
            )
            x_resources = uno.invoke(
                self._x_context.ServiceManager,
                "createInstanceWithArgumentsAndContext",
                (
                    "com.sun.star.resource.StringResourceWithLocation",
                    args,
                    self._x_context,
                ),
            )
            # Cache the resource
            self._resource_cache[dialog_name] = x_resources
        except Exception as ex:
            print(f"Error creating string resource for {dialog_name}: {ex}")

        return x_resources

    def translate(self, key, dialog_name="Strings"):
        """
        Get translated string from resource file

        Args:
            key: The property key to look up
            dialog_name: The dialog/resource file name (default: "Strings")

        Returns:
            Translated string if found, otherwise the key itself
        """
        x_resources = self.get_string_resource(dialog_name)

        if x_resources is not None:
            try:
                ids = x_resources.getResourceIDs()
                for resource_id in ids:
                    if key == resource_id:
                        return x_resources.resolveString(resource_id)
            except Exception as ex:
                print(f"Error translating key '{key}': {ex}")

        # Return key as fallback
        return key


# Module-level instance and function for easy access
_translator_instance = None


def translate(context, key, dialog_name="Strings"):
    """
    Module-level function to translate a string

    Args:
        context: UNO context
        key: The property key to look up
        dialog_name: The dialog/resource file name (default: "Strings")

    Returns:
        Translated string if found, otherwise the key itself

    Example:
        from translator import translate
        label = translate(ctx, "Symbol.000000")
    """
    global _translator_instance

    if _translator_instance is None:
        _translator_instance = Translator(context)

    return _translator_instance.translate(key, dialog_name)
