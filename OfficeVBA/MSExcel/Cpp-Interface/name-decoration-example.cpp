
/*print function name decoration(mangled) in windows*/

extern "C" void examplefunction(double arg1);

int main(void) {

       examplefunction(100);

}

 

void examplefunction(double arg1) {

       printf("Function Name: %s\n", __FUNCTION__);

       printf("Function Decorated Name %s\n", __FUNCDNAME__);

       printf("Function signature %s\n", __FUNCSIG__);

 

}
