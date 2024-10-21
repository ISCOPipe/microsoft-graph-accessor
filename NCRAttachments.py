from microsoft_graph_accessor import MicrosoftGraphAccessor
import asyncio
from msgraph.generated.sites.item.lists.item.items.item.list_item_item_request_builder import ListItemItemRequestBuilder
from kiota_abstractions.base_request_configuration import RequestConfiguration
from tqdm.asyncio import tqdm

class NCRAttachments(MicrosoftGraphAccessor):

    def __init__(self,credentials):
        super().__init__(credentials)
        self.get_NCR_library_location()

    def get_NCR_library_location(self):
        self.root = self.get_root()
        self.qc_site = self.get_site_from_root(self.root,"QC")
        self.ncr_list = self.get_list_from_site(self.qc_site,"NCR Log")

    # async def list_all_NCRs(self):
    #     ncrs = await self.get_all_list_items(self.qc_site.id, self.ncr_list.id)
    #     query_params = ListItemItemRequestBuilder.ListItemItemRequestBuilderGetQueryParameters(
    #             expand = ["fields"],
    #     )

    #     request_configuration = RequestConfiguration(
    #         query_parameters = query_params,
    #     )
    #     semaphore = asyncio.Semaphore(50)
    #     tasks = [self.get_list_item_fields(self.qc_site.id,self.ncr_list.id,ncr.id,semaphore,request_configuration) for ncr in ncrs]
    #     all_ncrs = []
    #     retry_tasks = []
    #     while True:
    #         try:
    #             ncr_details = await asyncio.gather(*tasks)
    #             all_ncrs.extend(ncr_details)
    #             break
    #         except Exception as e:
    #             if len(retry_tasks) == 0:
    #                 retry_tasks = tasks
    #             ncr_details = await asyncio.gather(*retry_tasks,return_exceptions=True)
    #             retry_tasks = [ncr for ncr in ncr_details if isinstance(ncr,Exception)]
    #             success_tasks = [ncr for ncr in ncr_details if not isinstance(ncr,Exception)]
    #             print(f"Successes: {len(success_tasks)}; Failures: {len(retry_tasks)}")
    #             all_ncrs.extend(success_tasks)
    #             if len(retry_tasks) == 0:
    #                 break
    #     return all_ncrs

    async def list_all_NCRs(self):
        ncrs = await self.get_all_list_items(self.qc_site.id, self.ncr_list.id)
        query_params = ListItemItemRequestBuilder.ListItemItemRequestBuilderGetQueryParameters(
            expand=["fields"],
        )

        request_configuration = RequestConfiguration(
            query_parameters=query_params,
        )
        
        semaphore = asyncio.Semaphore(50)  # Limits concurrent tasks
        tasks = [asyncio.gather(self.get_list_item_fields(self.qc_site.id, self.ncr_list.id, ncr.id, semaphore, request_configuration),self.get_list_item_attachments(self.qc_site.id,self.ncr_list.id,ncr.id,semaphore)) for ncr in ncrs[:1]]
        
        all_ncrs = []
        retry_tasks = tasks
        max_retries = 5
        retry_count = 0

        while retry_tasks and retry_count < max_retries:
            retry_count += 1
            print(f"Attempt {retry_count}: Processing {len(retry_tasks)} tasks...")
            
            ncr_details = await asyncio.gather(*tqdm(retry_tasks), return_exceptions=True)
            
            # Separate successful results and exceptions
            success_tasks = [ncr for ncr in ncr_details if not isinstance(ncr, Exception)]
            retry_tasks = [asyncio.gather(self.get_list_item_fields(self.qc_site.id, self.ncr_list.id, ncr.id, semaphore, request_configuration),self.get_list_item_attachments(self.qc_site.id,self.ncr_list.id,ncr.id,semaphore)) 
                        for ncr in ncr_details if isinstance(ncr, Exception)]

            print(f"Successes: {len(success_tasks)}; Failures: {len(retry_tasks)}")
            all_ncrs.extend(success_tasks)

            # If no tasks failed, we're done
            if not retry_tasks:
                break
            
            # Add a delay between retries (exponential backoff)
            await asyncio.sleep(2 ** retry_count)

        if retry_count == max_retries and retry_tasks:
            print(f"Failed to process {len(retry_tasks)} tasks after {max_retries} attempts.")
        
        return all_ncrs


    def get_ncr_attachments(self,ncr_id):
        ncr_attachments = self.get_drive_list()

        return ncr_attachments