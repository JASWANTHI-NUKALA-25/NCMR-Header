// A file is required to be in the root of the /src directory by the TypeScript compiler

export interface ITransmittalItem {
    Title: string;
    Modified: string;
    Created: string;
    EditorId: number;
    AuthorId: number;
}

export const TransmittalFieldNames = {
  Title:"Title"
}