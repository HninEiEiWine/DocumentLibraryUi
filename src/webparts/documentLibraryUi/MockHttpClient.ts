import { ISPList } from './DocumentLibraryUiWebPart';  
  
export default class MockHttpClient {  
    private static _items: ISPList[] = [,];  
    public static get(restUrl: string, options?: any): Promise<ISPList[]> {  
      return new Promise<ISPList[]>((resolve) => {  
            resolve(MockHttpClient._items);  
        });  
    }  
} 