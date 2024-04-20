def get_template_type(inverter_selection,structure_type,quotation_type,inverter_type,inverter2_type):
    template_file =""
    update_names_of_inverters_and_panels()
    if inverter_selection == "1":
        if inverter_type == "Grid Tie":
            if structure_type == "Normal":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGN_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGNSPI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNISPI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITN_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITNSPI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNISPI_Template.docx")
            if structure_type == "Raised":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGR_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (
                        ".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringNotIncluded"
                        "\\GTGRNI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (
                        ".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringNotIncluded\\GTGRNISPI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITR_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRNI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRSPI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRNISPI_Template.docx")
        if inverter_type == "Hybrid":
            if structure_type == "Normal":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGNSPI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNISPI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITN_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITNSPI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNISPI_Template.docx")
            if structure_type == "Raised":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGR_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGRSPI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNISPI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITR_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITRSPI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNISPI_Template.docx")
    if inverter_selection == "2":
        if inverter_type == "Grid Tie" and inverter2_type == "Hybrid":
            if structure_type == "Normal":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGN_WHI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGNSPI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNISPI_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITN_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNI_WHI_Template.docx")
                if quotation_type == "Itemised Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITNSPI_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNISPI_WHI_Template.docx")
            if structure_type == "Raised":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGR_WHI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNISPI_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITR_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTITRNI_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRSPI_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTITRNISPI_WHI_Template.docx")
        if inverter_type == "Grid Tie" and inverter2_type == "Grid Tie":
            if structure_type == "Normal":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGN_WGTI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTGNSPI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTGNNISPI_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITN_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal\\GridTieNormalNetMeteringIncluded"
                                     "\\GTITNSPI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieNormal"
                                     "\\GridTieNormalNetMeteringNotIncluded\\GTITNNISPI_WGTI_Template.docx")
            if structure_type == "Raised":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTR_WGTI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTGRSPI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTGRNISPI_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTIT_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTITRNI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised\\GridTieRaisedNetMeteringIncluded"
                                     "\\GTITRSPI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\GridTieInverters\\GridTieRaised"
                                     "\\GridTieRaisedNetMeteringNotIncluded\\GTITRNISPI_WGTI_Template.docx")
        if inverter_type == "Hybrid" and inverter2_type == "Grid Tie":
            if structure_type == "Normal":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_WGTI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGNSPI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNISPI_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITN_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITNSPI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNISPI_WGTI_Template.docx")
            if structure_type == "Raised":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGR_WGTI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGRSPI_WGTI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNISPI_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITR_WGTI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITRSPI_WGTI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNISPI_WGTI_Template.docx")
        if inverter_type == "Hybrid" and inverter2_type == "Hybrid":
            if structure_type == "Normal":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_WHI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HGN_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HGNNISPI_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITN_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNI_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringIncluded"
                                     "\\HITN_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridNormal\\HybridNormalNetMeteringNotIncluded"
                                     "\\HITNNISPI_WHI_Template.docx")
            if structure_type == "Raised":
                if quotation_type == "General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGR_WHI_Template.docx")
                if quotation_type == "General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HGRSPI_WHI_Template.docx")
                if quotation_type == "Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HGRNISPI_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITR_WHI_Template.docx")
                if quotation_type == "Itemised General Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNI_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringIncluded"
                                     "\\HITRSPI_WHI_Template.docx")
                if quotation_type == "Itemised Specify Brand Net Metering Not Included":
                    template_file = (".\\Templates\\HybridInverters\\HybridRaised\\HybridRaisedNetMeteringNotIncluded"
                                     "\\HITRNISPI_WHI_Template.docx")
    print(template_file)