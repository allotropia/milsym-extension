# SPDX-License-Identifier: MPL-2.0

PRJ = ../../..
SETTINGS = $(PRJ)/settings

include $(SETTINGS)/settings.mk
include $(SETTINGS)/std.mk

FILES = \
    Addons.xcu \
    META-INF/manifest.xml \
    description.xml \
    pkg-description/pkg-description.en \
    registration/license.txt \
    main.py

$(OUT_BIN)/minimal-python.$(UNOOXT_EXT): $(FILES)
	-$(MKDIR) $(subst /,$(PS),$(@D))
	$(SDK_ZIP) $@ $^
