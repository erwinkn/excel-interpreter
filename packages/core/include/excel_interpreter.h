#ifndef EXCEL_INTERPRETER_H
#define EXCEL_INTERPRETER_H

#include <stddef.h>
#include <stdint.h>

#ifdef __cplusplus
extern "C" {
#endif

typedef enum ei_status {
    EI_STATUS_OK = 0,
    EI_STATUS_INVALID_ARGUMENT = 1,
    EI_STATUS_UNSUPPORTED_FORMULA = 2,
} ei_status;

typedef enum ei_value_tag {
    EI_VALUE_BLANK = 0,
    EI_VALUE_NUMBER = 1,
    EI_VALUE_BOOLEAN = 2,
} ei_value_tag;

typedef struct ei_value {
    int32_t tag;
    double number;
    uint8_t boolean;
} ei_value;

uint32_t ei_version_major(void);
uint32_t ei_version_minor(void);
uint32_t ei_version_patch(void);
const char *ei_version_string(void);
double ei_add_f64(double lhs, double rhs);
const char *ei_demo_greeting(void);
ei_status ei_eval_formula_utf8(const char *formula_ptr, size_t formula_len, ei_value *out_value);

#ifdef __cplusplus
}
#endif

#endif
