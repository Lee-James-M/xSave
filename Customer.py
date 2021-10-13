
class Customer:

    def __init__(self):
        print('Customer Class Created')

    @staticmethod
    def get_customer_name():
        with open("C:\\pcdmisw\\pcddata\\isocust.dat", "r") as f:
            f.readline()
            customer_name = f.readline()
            customer_name = customer_name.rstrip("\n")
            print("customer name is " + customer_name)
            return customer_name

    @staticmethod
    def get_customer_address():
        with open("C:\\pcdmisw\\pcddata\\isocust.dat", "r") as f:
            f.readline()
            f.readline()
            customer_address = f.readline()
            customer_address = customer_address.rstrip("\n")
            print("customer address is " + customer_address)
            return customer_address

    @staticmethod
    def get_customer_machine_serial():
        with open("C:\\pcdmisw\\pcddata\\isomac.dat", "r") as f:
            for i in range(1, 6):
                mac_serial_no = f.readline()
                mac_serial_no = mac_serial_no.rstrip("\n")
            return mac_serial_no

    @staticmethod
    def get_customer_machine_size():
        with open("C:\\pcdmisw\\pcddata\\isomac.dat", "r") as f:
            for i in range(1, 5):
                mac_size = f.readline()
                mac_size = mac_size.rstrip("\n")
            print(f'Customer machine size is: {mac_size}')
            return mac_size

    @staticmethod
    def get_customer_machine_model():
        with open("C:\\pcdmisw\\pcddata\\isomac.dat", "r") as f:
            for i in range(1, 3):
                mac_model = f.readline()
                mac_model = mac_model.rstrip("\n")
            return mac_model

    @staticmethod
    def get_customer_machine_spec():
        pass

    @staticmethod
    def customer_info():
        with open("C:\\pcdmisw\\pcddata\\isocust.dat", "r") as f:
            f.readline()
            custdat1 = f.readline()
            custdat1 = custdat1.rstrip("\n")
            custdat2 = f.readline()
            custdat2 = custdat2.rstrip("\n")
            cust_info = custdat1 + ", " + custdat2
            print("customer info is " + cust_info)
            return cust_info

    @staticmethod
    def get_isomac_dat_file_data():
        with open("C:\\pcdmisw\\pcddata\\isomac.dat", "r") as f:
            isomac = f.readlines()
            print("em1 is below - isomac.dat data in a list")
            print(isomac)
        return isomac
