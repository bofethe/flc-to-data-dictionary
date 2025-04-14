from arcgis.gis import GIS
from arcgis.features import FeatureLayerCollection
import pandas as pd
import datetime
import re
import os

now = datetime.datetime.now().strftime("%Y%m%d%H%M")

# Connect to AGOL
gis = GIS("home")

def flc_to_data_dict(id, dir=None):
    '''
    This function takes an item ID of a Feature Layer Collection 
    and creates an Excel file with the data dictionary for each layer 
    in the collection.

    :param int id: Item ID of the Feature Layer Collection.
    :param str (Optional) dir: Directory to save the Excel file. If not provided,
                                the file will be saved in the current working directory.
    '''
    item = gis.content.get(id)
    if dir:
        fpath = os.path.join(dir, f"{item.title}_DataDictionary_{now}.xlsx")
    else:
        fpath = f"{item.title}_DataDictionary_{now}.xlsx"
    flc = FeatureLayerCollection.fromitem(item)

    # Get each layer's field information
    with pd.ExcelWriter(fpath) as writer:
        for layer in flc.layers:
            layer_name = layer.properties.name or f"Layer_{layer.properties.id}"
            fields = layer.properties.fields

            field_info = []
            for field in fields:

                # get domain values if they exist
                domain_str = None
                domain = field.get("domain")
                if domain:
                    if domain.get("type") == "codedValue":
                        domain_str = "; ".join(f"{cv['code']}: {cv['name']}" for cv in domain.get("codedValues", []))
                    elif domain.get("type") == "range":
                        domain_str = f"Range: {domain.get('minValue')} to {domain.get('maxValue')}"

                field_info.append({
                    "Field Name": field['name'],
                    "Alias": field.get('alias', None),
                    "Field Type": field['type'],
                    "Length": field.get('length', None),
                    "Nullable": field.get('nullable', None),
                    "Editable": field.get('editable', None),
                    "Domain Name": field.get('domain', {}).get('name') if field.get('domain') else None,
                    "Domain Values": domain_str,
                    "Default Value": field.get('defaultValue', None)
                })

            # clean field names and write to Excel
            df = pd.DataFrame(field_info)
            clean_name = re.sub(r'[^A-Za-z0-9 _]', '', layer_name)[:31] #xlsx max length for sheet name
            df.to_excel(writer, sheet_name=clean_name, index=False)

flc_to_data_dict("0e28ef312008491aa86f90bd9ca7c706") #NOTE Replace with your item ID