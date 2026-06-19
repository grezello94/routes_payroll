export interface EmployeeVerificationRecord {
  id?: number;
  companyId: number;
  primaryBioData: {
    fullName: string;
    dateOfBirth: string;
    age: number | null;
    currentAddress: string;
    permanentAddress: string;
  };
  mediaDocuments: {
    photoDataUrl: string;
    idProofs: Array<{
      type: string;
      reference: string;
      fileName: string;
      fileDataUrl: string;
    }>;
  };
  medical: {
    enabled: boolean;
    fitnessStatus: "fit" | "review" | "unfit" | "";
    certificateFileName: string;
    certificateDataUrl: string;
  };
  internalScreening: {
    classification: "sensitive/internal_only";
    consumesTobacco: boolean;
    consumesAlcohol: boolean;
    substanceAbuseSigns: boolean;
    zeroToleranceAcknowledged: boolean;
  };
  emergencyVerification: {
    emergencyContactName: string;
    relationship: string;
    phoneNumber: string;
    consentToPoliceVerification: boolean;
  };
  createdAt?: string;
  updatedAt?: string;
}
