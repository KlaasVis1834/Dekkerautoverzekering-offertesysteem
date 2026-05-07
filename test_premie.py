from cockpit_premie import haal_premie_uit_cockpit

url = "https://qis.asrcockpit.nl/Axon/add-quotation/form/fill-form/screen/S3?formExtId=autoZakelijk&formDialogueExtId=offerte&formBehaviourExtId=%257B%2522prefix%2522%253A%2522A%2522%252C%2522agreementExtId%2522%253A%25221109%2522%257D&dateFormDefinition=2026-02-03&originExtId=Extranet"

r = haal_premie_uit_cockpit(url, storage_state_path="state.json", headless=True)

print("landed:", r.landed_url)
print("premie:", r.premie)
print("assurantiebelasting:", r.assurantiebelasting)
print("eindpremie_excl:", r.eindpremie_excl)
print("raw_keys_count:", len(r.raw_fields))
