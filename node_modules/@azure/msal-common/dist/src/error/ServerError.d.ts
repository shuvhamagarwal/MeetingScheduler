import { AuthError } from "./AuthError";
/**
 * Error thrown when there is an error with the server code, for example, unavailability.
 */
export declare class ServerError extends AuthError {
    subError: string;
    constructor(errorCode: string, errorMessage?: string, subError?: string);
}
