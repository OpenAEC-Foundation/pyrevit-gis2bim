# -*- coding: utf-8 -*-
"""GIS2BIM API clients."""

from .pdok import (
    # Locatieserver
    PDOKLocatie,
    LocationData,
    # BGT
    PDOKBGT,
    # WMTS
    PDOKWMTS,
    # Kadastrale Kaart
    PDOKKadaster,
    PerceelData,
    PerceelAnnotatie,
    OpenbareRuimteNaam,
)

from .wfs import (
    WFSClient,
    WFSLayer,
    WFSFeature,
    get_wfs_data,
)

from .wfs_layers import (
    PDOK_LAYERS,
    DEFAULT_LAYERS,
    get_layer,
    get_all_layers,
    get_active_layers,
    get_default_layers,
    get_layers_by_category,
)

from .ogc_api import (
    OGCAPIClient,
    OGCAPICollection,
    OGCAPIFeature,
)

from .bgt_layers import (
    BGT_API_URL,
    BGT_LAYERS,
    BGT_VEGETATIEOBJECT,
    BGT_PAAL,
    get_bgt_layer,
    get_all_bgt_layers,
    get_active_bgt_layers,
    get_default_bgt_layers,
    get_bgt_layers_by_category,
)

from .ahn import (
    AHNClient,
    AHNError,
)

from .bag3d import (
    BAG3DClient,
    BAG3DError,
    BAG3DTile,
)

from .wms import (
    WMSClient,
    WMS_LAYERS,
    WMS_CATEGORIES,
    get_layers_by_category as get_wms_layers_by_category,
    get_layer as get_wms_layer,
)

from .streetview import (
    StreetViewClient,
)

from .wmts_tiles import (
    ArcGISTileClient,
    TIJDREIS_YEARS,
)

from .nap import (
    NAPClient,
    NAPPeilmerk,
)

from .natura2000 import (
    Natura2000Client,
    Natura2000Area,
    Natura2000Result,
)

from .bro import (
    BROClient,
    CPTData,
    BHRData,
    Grondlaag,
    GRONDSOORT_KLEUREN,
    CPT_KLEUR,
    get_grondsoort_kleur,
    classificeer_grondsoort,
)
