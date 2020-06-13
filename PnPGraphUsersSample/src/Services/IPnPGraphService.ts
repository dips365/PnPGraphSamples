export interface IPnPGraphService{
  getCurrentUser():Promise<any[]>;
  getMatchingUser(email:string):Promise<any[]>;
  getPeopleAroundMe():Promise<any[]>;
}
