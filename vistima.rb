require "roo"
require "pastel"
require "byebug"
require "tty-spinner"

def print_vistima
    pastel = Pastel.new
    format = "[#{pastel.yellow(":spinner")}] " + pastel.cyan("Looking for vistima =>")
    spinner_ok = TTY::Spinner.new(format, success_mark: pastel.green("+"))
    spinner_error = TTY::Spinner.new(format, success_mark: pastel.red("x"))
    
    begin 
        file = Roo::Excelx.new("./estudiantes.xlsx")
        if valid_arguments?(file)
            20.times do
                spinner_ok.spin
                sleep(0.1)
            end
            spinner_ok.success(pastel.green(get_vistima(file)))
        else 
            spinner_error.success(pastel.red('Group not allowed'))
        end
    rescue => e
        spinner_error.success(pastel.red(e))
    end
end

private

def valid_arguments?(file)
    valid_groups = file.parse(headers: true).first.keys
    if valid_groups.include?(ARGV[0].upcase) 
        true
    else
        pastel = Pastel.new
        puts pastel.cyan("Groups available => ") + pastel.magenta(valid_groups.join(', '))
        false
    end 
end

def get_vistima(file)
    (file.first_column..file.first_column).each do |column_number|
        if file.column(column_number).first.eql?(ARGV[0].upcase)
            return file.column(column_number).drop(1).sample
        end
    end
end


print_vistima()