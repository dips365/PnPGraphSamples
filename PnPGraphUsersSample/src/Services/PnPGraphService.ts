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

  public async getPeopleAroundMe():Promise<any[]>{
    let peopleinfo:any[] = [];
    try {
      await graph.me.people().then((info)=>{
        peopleinfo = info
      }).catch((error)=>{
        Log.error("",error);
      })
    } catch (error) {
      Log.error("",error);
    }
    return peopleinfo;
  }

  public async getMatchingUser(email:string):Promise<any[]>{
    let info:any[] = [];
    try {
      await graph.users.getById(email).get().then((info)=>{
        info = info;
      });
    } catch (error) {
      Log.error("",error);
    }
    return info;
  }
}
