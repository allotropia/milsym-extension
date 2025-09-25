# SPDX-License-Identifier: MPL-2.0

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
