import { Request, Response, NextFunction } from 'express';
import { InspectionFormSchema } from '../validation/schemas';
import { ApiResponse } from '../types';

export const validateInspectionForm = (
  req: Request,
  res: Response<ApiResponse>,
  next: NextFunction
) => {
  try {
    const validatedData = InspectionFormSchema.parse(req.body);
    req.body = validatedData;
    next();
  } catch (error: any) {
    console.error('Validation error:', error);
    
    const errorMessage = error.errors 
      ? error.errors.map((err: any) => `${err.path.join('.')}: ${err.message}`).join(', ')
      : 'Invalid form data';
    
    res.status(400).json({
      success: false,
      error: 'Validation failed',
      message: errorMessage
    });
  }
};

export const errorHandler = (
  error: Error,
  _req: Request,
  res: Response<ApiResponse>,
  _next: NextFunction
) => {
  console.error('Server error:', error);
  
  res.status(500).json({
    success: false,
    error: 'Internal server error',
    message: error.message
  });
};