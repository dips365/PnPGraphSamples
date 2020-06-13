import { IPnPGraphService } from "./IPnPGraphService";
import { Log } from "@microsoft/sp-core-library";
import { graph } from "@pnp/graph";
import "@pnp/graph/users";
export class PnPGraphService implements IPnPGraphService{
  public async getCurrentUser():Promise<any[]>{
    let myInfo:any[] = [];
    try {
      const currentUser = await graph.me();
      if(currentUser){
        alert(currentUser);
      }
    } catch (error) {
      Log.error("",error);
    }

    return myInfo;
  }


}
