package com.excel.exception;

import lombok.experimental.StandardException;
import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.ResponseStatus;

@ResponseStatus(value = HttpStatus.BAD_REQUEST)
@StandardException
public class EmptyExcelFileException extends RuntimeException{

  public EmptyExcelFileException(String message) {
    super(message);
  }

}
