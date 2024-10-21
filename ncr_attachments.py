from NCRAttachments import NCRAttachments
from functools import lru_cache
import time
import asyncio

CACHE_EXPIRY = 3600*24  # 1 day

@lru_cache(maxsize=1)
def get_cached_secrets(timestamp,secret_names=()):
    # This function will only be called when the timestamp changes
    # Implement your chosen method of secret retrieval here
    # For this example, we'll use AWS Systems Manager Parameter Store
    import boto3
    session = boto3.Session(region_name='us-east-2')
    ssm = session.client('ssm')
    response = ssm.get_parameters(
        Names=secret_names,
        WithDecryption=True
    )
    return {param['Name']: param['Value'] for param in response['Parameters']}

def get_secrets(secret_names):
    # Round the current timestamp to the nearest hour
    current_time = int(time.time() / CACHE_EXPIRY) * CACHE_EXPIRY
    return get_cached_secrets(current_time,secret_names=secret_names)

secrets = get_secrets(secret_names=('SHAREPOINT_TENANT_ID','SHAREPOINT_CLIENT_ID','SHAREPOINT_CLIENT_SECRET'))
    
if __name__ == "__main__":
    credentials={
            'tenant_name_or_id': secrets.get('SHAREPOINT_TENANT_ID'),
            'client_id': secrets.get('SHAREPOINT_CLIENT_ID'),
            'client_secret': secrets.get('SHAREPOINT_CLIENT_SECRET'),
    }
    ncr_attachments = NCRAttachments(credentials)
    #print(ncr_attachments.get_list_operations(ncr_attachments.ncr_list))
    ncrs = asyncio.run(ncr_attachments.list_NCRs())
    print(ncrs[0])
    print(len(ncrs))
    print(ncrs[-1].__dict__)
    #ncr_attachments.get_ncr_attachments("1234")