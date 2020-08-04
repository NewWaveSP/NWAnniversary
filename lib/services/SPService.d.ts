import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ApplicationCustomizerContext } from "@microsoft/sp-application-base";
export declare class SPService {
    private _context;
    private graphClient;
    constructor(_context: WebPartContext | ApplicationCustomizerContext);
    getPAnniversarys(upcommingDays: number): Promise<any[]>;
}
export default SPService;
